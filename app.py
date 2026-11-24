"""ASD scanpath screening app.

Pipeline (notebook-exact when weights are present):
  preprocess_scanpath → MobileNet (1024-D, frozen) → PCA (250) →
  StackingClassifier (SVM poly degree 4 + KNN k=3 → logistic regression)

Class codes match the training notebook: 1 = ASD, 0 = Non-ASD.

If pca_model.pkl and asd_stacking_model.pkl are missing, or TensorFlow /
scikit-learn cannot load them, the app stays up and scores the image with
contour geometry so the interface can be tested. That path is labeled
simulation and is not a model prediction.
"""

from __future__ import annotations

import os

import cv2
import numpy as np
from flask import Flask, jsonify, render_template, request

APP_ROOT = os.path.dirname(os.path.abspath(__file__))
PCA_PATH = os.path.join(APP_ROOT, "pca_model.pkl")
MODEL_PATH = os.path.join(APP_ROOT, "asd_stacking_model.pkl")

IMG_SIZE = (224, 224)
ALLOWED_EXTENSIONS = {"png", "jpg", "jpeg"}
NORM_CENTRAL_SALIENCE = 0.68

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 8 * 1024 * 1024

feature_extractor = None
mobilenet_preprocess = None
pca_model = None
stacking_model = None
pipeline_mode = "simulation"
pipeline_notice = (
    "Trained weights were not loaded. Place pca_model.pkl and "
    "asd_stacking_model.pkl next to app.py to run the MobileNet stacking pipeline. "
    "Current scores use contour geometry only and are not model predictions."
)


def preprocess_scanpath(img_bgr):
    img = cv2.cvtColor(img_bgr, cv2.COLOR_BGR2RGB)
    img = cv2.resize(img, IMG_SIZE)
    img = cv2.GaussianBlur(img, (5, 5), 0)  # Light Gaussian noise filtering
    img = img.astype(np.float32)
    img = (img - img.min()) / (img.max() - img.min() + 1e-8)  # Min-Max normalize [0, 1]
    return img


def _load_pipeline():
    """Load MobileNet + PCA + stacking weights. Never raise to the caller."""
    global feature_extractor, mobilenet_preprocess, pca_model, stacking_model
    global pipeline_mode, pipeline_notice

    if not (os.path.isfile(PCA_PATH) and os.path.isfile(MODEL_PATH)):
        pipeline_mode = "simulation"
        pipeline_notice = (
            "pca_model.pkl and asd_stacking_model.pkl were not found in the app folder. "
            "Analysis is running in contour-simulation mode so the interface can be tested. "
            "These scores are not predictions from the trained stacking classifier."
        )
        return

    try:
        import joblib
        from tensorflow.keras.applications import MobileNet
        from tensorflow.keras.applications.mobilenet import preprocess_input as _mobilenet_preprocess
        from tensorflow.keras.models import Model

        base = MobileNet(
            weights="imagenet",
            include_top=False,
            input_shape=(224, 224, 3),
            pooling="avg",
        )
        base.trainable = False
        feature_extractor = Model(inputs=base.input, outputs=base.output)
        mobilenet_preprocess = _mobilenet_preprocess
        pca_model = joblib.load(PCA_PATH)
        stacking_model = joblib.load(MODEL_PATH)
        pipeline_mode = "model"
        pipeline_notice = None
    except Exception as exc:
        feature_extractor = None
        mobilenet_preprocess = None
        pca_model = None
        stacking_model = None
        pipeline_mode = "simulation"
        pipeline_notice = (
            "The saved PCA or stacking model could not be loaded "
            f"({exc.__class__.__name__}). Analysis fell back to contour simulation "
            "and these scores are not model predictions."
        )


def _model_asd_probability(img_norm: np.ndarray) -> float:
    """1024-D MobileNet vector → 250 PCA components → P(ASD)."""
    batch_uint8 = (np.expand_dims(img_norm, axis=0) * 255.0).astype(np.uint8)
    batch_prep = mobilenet_preprocess(batch_uint8.astype(np.float32))
    features = feature_extractor.predict(batch_prep, verbose=0)
    reduced = pca_model.transform(features)

    if hasattr(stacking_model, "predict_proba"):
        proba = stacking_model.predict_proba(reduced)[0]
        classes = list(getattr(stacking_model, "classes_", [0, 1]))
        index = classes.index(1) if 1 in classes else len(classes) - 1
        return float(proba[index])

    predicted = int(stacking_model.predict(reduced)[0])
    return 0.85 if predicted == 1 else 0.15


def _trace_mask(img_norm: np.ndarray) -> np.ndarray:
    """Binary mask of scanpath ink, whether the trace is light-on-dark or dark-on-light."""
    gray = cv2.cvtColor((img_norm * 255.0).astype(np.uint8), cv2.COLOR_RGB2GRAY)
    blurred = cv2.GaussianBlur(gray, (5, 5), 0)
    gx = cv2.Sobel(blurred, cv2.CV_32F, 1, 0, ksize=3)
    gy = cv2.Sobel(blurred, cv2.CV_32F, 0, 1, ksize=3)
    magnitude = cv2.magnitude(gx, gy)
    cutoff = float(np.percentile(magnitude, 90))
    mask = (magnitude >= max(cutoff, 1.0)).astype(np.uint8)

    # Contours catch thin strokes the percentile gate can fragment.
    edges = cv2.Canny(blurred, 40, 120)
    mask = cv2.bitwise_or(mask * 255, edges)
    kernel = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (3, 3))
    mask = cv2.morphologyEx(mask, cv2.MORPH_CLOSE, kernel)
    return mask


def _path_tortuosity(points: list) -> float:
    """Path length between fixation blobs divided by the span from first to last."""
    if len(points) < 2:
        return 1.0
    remaining = [np.asarray(point, dtype=np.float64) for point in points[1:]]
    path = [np.asarray(points[0], dtype=np.float64)]
    while remaining:
        last = path[-1]
        distances = [float(np.linalg.norm(point - last)) for point in remaining]
        nearest = int(np.argmin(distances))
        path.append(remaining.pop(nearest))
    path_arr = np.vstack(path)
    segment_length = float(np.linalg.norm(np.diff(path_arr, axis=0), axis=1).sum())
    span = float(np.linalg.norm(path_arr[-1] - path_arr[0])) + 1e-6
    return float(np.clip(segment_length / span, 1.0, 6.0))


def analyze_scanpath_geometry(img_norm: np.ndarray) -> dict:
    """Quantify fixation clumping, central salience, tortuosity, and exploration.

    These are geometric proxies measured on the preprocessed scanpath. They feed
    the reasoning cards in both model and simulation modes. Simulation mode also
    uses them as the classifier, because no trained weights are available.
    """
    mask = _trace_mask(img_norm)
    trace = mask > 0
    height, width = trace.shape
    total = int(trace.sum())
    area = float(height * width)

    x0, x1 = int(width * 0.30), int(width * 0.70)
    y0, y1 = int(height * 0.30), int(height * 0.70)

    if total < 25:
        return {
            "fixation_dwell_ratio": 0.50,
            "central_salience": 0.50,
            "scanpath_tortuosity": 1.00,
            "visual_entropy": 0.50,
            "trace_fraction": total / area,
            "signal_quality": "low",
        }

    central = int(trace[y0:y1, x0:x1].sum())
    central_salience = central / total
    # High when gaze mass sits outside the central anchors (peripheral dwell).
    fixation_dwell_ratio = float(np.clip(1.0 - central_salience, 0.0, 1.0))

    grid = 8
    bins = np.zeros((grid, grid), dtype=np.float64)
    ys, xs = np.where(trace)
    gi = np.clip((ys * grid) // height, 0, grid - 1)
    gj = np.clip((xs * grid) // width, 0, grid - 1)
    np.add.at(bins, (gi, gj), 1.0)
    occupied = bins[bins > 0]
    probs = occupied / occupied.sum()
    entropy = float(-(probs * np.log2(probs)).sum() / np.log2(grid * grid))

    kernel = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (11, 11))
    blobs = cv2.dilate((trace.astype(np.uint8) * 255), kernel, iterations=1)
    count, _labels, stats, centroids = cv2.connectedComponentsWithStats(blobs, connectivity=8)
    points = []
    for index in range(1, count):
        if stats[index, cv2.CC_STAT_AREA] < 40:
            continue
        points.append(centroids[index].astype(np.float64))
    tortuosity = _path_tortuosity(points)

    return {
        "fixation_dwell_ratio": fixation_dwell_ratio,
        "central_salience": float(central_salience),
        "scanpath_tortuosity": tortuosity,
        "visual_entropy": entropy,
        "trace_fraction": total / area,
        "signal_quality": "ok",
    }


def _simulation_asd_probability(metrics: dict) -> float:
    """Map contour markers onto a probability. Deterministic for a given image."""
    if metrics["signal_quality"] == "low":
        return 0.50

    clump = np.clip((metrics["fixation_dwell_ratio"] - 0.28) / 0.42, 0.0, 1.0)
    peripheral = np.clip((0.62 - metrics["central_salience"]) / 0.40, 0.0, 1.0)
    jitter = np.clip((metrics["scanpath_tortuosity"] - 1.15) / 1.50, 0.0, 1.0)
    restricted = np.clip((0.72 - metrics["visual_entropy"]) / 0.40, 0.0, 1.0)
    score = 0.34 * clump + 0.30 * peripheral + 0.20 * jitter + 0.16 * restricted
    return float(np.clip(0.08 + 0.86 * score, 0.08, 0.94))


def _severity(level: str) -> str:
    if level not in {"low", "medium", "high"}:
        return "medium"
    return level


def _format_metrics(metrics: dict) -> dict:
    dwell = metrics["fixation_dwell_ratio"]
    tort = metrics["scanpath_tortuosity"]
    entropy = metrics["visual_entropy"]
    dwell_tag = "Atypical" if dwell >= 0.55 else "Typical"
    if tort >= 1.85:
        tort_tag = "High"
    elif tort >= 1.35:
        tort_tag = "Moderate"
    else:
        tort_tag = "Low"
    if entropy < 0.45:
        entropy_tag = "Low exploration"
    elif entropy < 0.70:
        entropy_tag = "Moderate exploration"
    else:
        entropy_tag = "Broad exploration"
    return {
        "fixation_dwell_ratio": f"{dwell * 100:.0f}% ({dwell_tag})",
        "scanpath_tortuosity": f"{tort:.2f} ({tort_tag})",
        "visual_entropy": f"{entropy:.2f} ({entropy_tag})",
    }


def build_reasoning(has_asd: bool, metrics: dict) -> dict:
    """Clinical-style factor cards grounded in the measured geometry."""
    dwell = metrics["fixation_dwell_ratio"]
    central = metrics["central_salience"]
    tort = metrics["scanpath_tortuosity"]
    entropy = metrics["visual_entropy"]
    central_pct = int(round(central * 100))
    norm_pct = int(round(NORM_CENTRAL_SALIENCE * 100))

    if dwell >= 0.62:
        fixation_status, fixation_severity = "High Divergence", "high"
    elif dwell >= 0.45:
        fixation_status, fixation_severity = "Moderate Clustering", "medium"
    else:
        fixation_status, fixation_severity = "Distributed", "low"

    if central < 0.45:
        salience_status, salience_severity = f"Reduced ({central_pct}% vs {norm_pct}% norm)", "high"
    elif central < NORM_CENTRAL_SALIENCE:
        salience_status, salience_severity = f"Below norm ({central_pct}% vs {norm_pct}% norm)", "medium"
    else:
        salience_status, salience_severity = f"Preserved ({central_pct}% vs {norm_pct}% norm)", "low"

    if tort >= 1.85:
        saccade_status, saccade_severity = "Elevated Jitter", "high"
    elif tort >= 1.35:
        saccade_status, saccade_severity = "Mild Irregularity", "medium"
    else:
        saccade_status, saccade_severity = "Regular Transitions", "low"

    if has_asd:
        fixation_desc = (
            "Fixations are clustered in isolated peripheral zones rather than "
            "exploring holistic scene features."
            if dwell >= 0.45 or central < 0.55
            else "The classifier indicates an ASD-like representation. Fixation mass is only mildly concentrated on this image."
        )
        salience_desc = (
            "Gaze is diverted away from central primary anchors, with reduced dwell on socially salient structure."
            if central < NORM_CENTRAL_SALIENCE
            else "Central anchor dwell is near the reference range. The ASD indication is carried more by the learned feature stack than by this salience proxy."
        )
        saccade_desc = (
            "Shorter, irregular saccade jumps indicate fragmented visual navigation."
            if tort >= 1.35 or entropy < 0.55
            else "Saccade geometry is relatively smooth. Fragmentation is not the dominant marker on this scanpath."
        )
        summary = (
            "The patient's visual scanpath demonstrates significant peripheral fixation "
            "clumping and reduced exploration of central visual anchors."
        )
        if central >= 0.60 and entropy >= 0.65 and dwell < 0.45:
            summary = (
                "The model indicates an ASD-like scanpath pattern. Geometric markers show "
                "more central exploration than a classic peripheral-clumping profile, so "
                "this result should be read with the raw trace and clinical history."
            )
        recommendation = "Recommend comprehensive clinical evaluation with ADOS-2 protocol."
        label = "Autism Spectrum Disorder (ASD) Pattern Indicated"
    else:
        fixation_desc = (
            "Fixation mass is spread across the scene, consistent with holistic visual exploration."
            if dwell < 0.55
            else "Some fixation clustering is present, but the overall classifier call remains neurotypical."
        )
        salience_desc = (
            "Gaze connects key focal points, including central anchors, in a stable triadic pattern."
            if central >= 0.50
            else "Central salience is lower than the reference range. The classifier still favors a neurotypical pattern."
        )
        saccade_desc = (
            "Scanning paths are comparatively smooth, with balanced jump lengths between fixations."
            if tort < 1.85
            else "Path irregularity is elevated. It is recorded here, but it did not overturn the neurotypical classifier call."
        )
        summary = (
            "The scanpath shows holistic visual exploration, with smooth scanning paths "
            "and a balanced fixation distribution across primary focal points."
        )
        recommendation = (
            "Geometric and classifier findings are closer to a neurotypical exploration pattern. "
            "This is a screening aid only and is not clinical clearance."
        )
        label = "Neurotypical Visual Pattern"

    if metrics["signal_quality"] == "low":
        summary = (
            "Very little scanpath structure was detected in this image. "
            "The score is uncertain and should not be interpreted as a screening result."
        )
        recommendation = "Upload a scanpath rendering with a visible gaze trace and repeat the analysis."

    return {
        "prediction_label": label,
        "clinical_summary": summary,
        "explainability": {
            "key_factors": [
                {
                    "title": "Atypical Fixation Dispersion" if has_asd else "Holistic Visual Exploration",
                    "status": fixation_status,
                    "severity": _severity(fixation_severity),
                    "description": fixation_desc,
                },
                {
                    "title": "Social-Salience Avoidance" if has_asd else "Stable Triadic Gaze",
                    "status": salience_status,
                    "severity": _severity(salience_severity),
                    "description": salience_desc,
                },
                {
                    "title": "Saccadic Path Fragmentation" if has_asd else "Saccadic Path Continuity",
                    "status": saccade_status,
                    "severity": _severity(saccade_severity),
                    "description": saccade_desc,
                },
            ],
            "metrics": _format_metrics(metrics),
            "recommendation": recommendation,
        },
    }


def analyze_image(img_bgr: np.ndarray) -> dict:
    img_norm = preprocess_scanpath(img_bgr)
    metrics = analyze_scanpath_geometry(img_norm)

    mode = pipeline_mode
    if mode == "model" and feature_extractor is not None:
        try:
            asd_probability = _model_asd_probability(img_norm)
        except Exception as exc:
            mode = "simulation"
            asd_probability = _simulation_asd_probability(metrics)
            notice = (
                "Model inference failed "
                f"({exc.__class__.__name__}). Contour simulation was used instead "
                "and this score is not a stacking-classifier prediction."
            )
        else:
            notice = None
    else:
        asd_probability = _simulation_asd_probability(metrics)
        notice = pipeline_notice

    asd_probability = float(np.clip(asd_probability, 0.0, 1.0))
    has_asd = bool(asd_probability >= 0.50)
    confidence = asd_probability if has_asd else (1.0 - asd_probability)
    reasoning = build_reasoning(has_asd, metrics)

    payload = {
        "has_asd": has_asd,
        "prediction_label": reasoning["prediction_label"],
        "probability": round(asd_probability, 3),
        "confidence": round(confidence * 100, 1),
        "clinical_summary": reasoning["clinical_summary"],
        "explainability": reasoning["explainability"],
        "pipeline_mode": mode,
    }
    if notice:
        payload["notice"] = notice
    return payload


def _allowed(filename: str) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/upload", methods=["POST"])
def upload():
    uploaded = request.files.get("file")
    if uploaded is None or uploaded.filename == "":
        return jsonify({"error": "Choose a PNG or JPG scanpath image."}), 400
    if not _allowed(uploaded.filename):
        return jsonify({"error": "Only PNG and JPG images are accepted."}), 400

    raw = uploaded.read()
    if not raw:
        return jsonify({"error": "The uploaded file is empty."}), 400

    encoded = np.frombuffer(raw, dtype=np.uint8)
    image = cv2.imdecode(encoded, cv2.IMREAD_COLOR)
    if image is None:
        return jsonify({"error": "That file could not be read as an image."}), 400

    return jsonify(analyze_image(image))


@app.errorhandler(413)
def too_large(_error):
    return jsonify({"error": "Image must be smaller than 8 MB."}), 413


_load_pipeline()


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5050))
    app.run(debug=True, host="0.0.0.0", port=port)
