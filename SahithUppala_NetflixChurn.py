# -*- coding: utf-8 -*-
"""
=============================================================================
  Netflix Customer Churn & Subscriber Retention Prediction
  Author  : Sahith Uppala
  Dataset : netflix_customer_churn.csv  (5 000 records, 14 features)
  Purpose : End-to-end ML pipeline + executive 4-panel dashboard
=============================================================================
  Sections
  --------
  1. Imports & configuration
  2. Data loading & exploration
  3. Feature engineering & pre-processing
  4. Model training  - Random Forest (primary) + Logistic Regression (baseline)
  5. Model evaluation - accuracy, precision, recall, F1, confusion matrix
  6. Executive 4-panel dashboard  -> saved as netflix_executive_dashboard.png
  7. Console report summary
=============================================================================
"""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

# -----------------------------------------------------------------------------
# 1. IMPORTS & CONFIGURATION
# -----------------------------------------------------------------------------
import warnings
warnings.filterwarnings("ignore")

import numpy  as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")                          # headless - no GUI window needed
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import seaborn as sns

from sklearn.model_selection   import train_test_split, StratifiedKFold, cross_val_score
from sklearn.preprocessing     import LabelEncoder, StandardScaler
from sklearn.ensemble          import RandomForestClassifier
from sklearn.linear_model      import LogisticRegression
from sklearn.metrics           import (accuracy_score, precision_score,
                                       recall_score, f1_score,
                                       confusion_matrix, classification_report,
                                       roc_auc_score, roc_curve)
from sklearn.pipeline          import Pipeline

# Plotting aesthetics
NETFLIX_RED   = "#E50914"
NETFLIX_DARK  = "#141414"
NETFLIX_GREY  = "#564D4D"
ACCENT        = "#B20710"
PALETTE       = [NETFLIX_RED, "#AAAAAA"]

sns.set_theme(style="darkgrid", palette="muted", font_scale=1.05)
DASHBOARD_PATH = "netflix_executive_dashboard.png"

# -----------------------------------------------------------------------------
# 2. DATA LOADING & EXPLORATION
# -----------------------------------------------------------------------------
print("=" * 70)
print("  Netflix Customer Churn & Retention Prediction")
print("=" * 70)

df = pd.read_csv("netflix_customer_churn.csv")
print(f"\n[DATA]  Shape          : {df.shape}")
print(f"[DATA]  Columns        : {list(df.columns)}")
print(f"[DATA]  Missing values : {df.isnull().sum().sum()}")
print(f"\n[DATA]  Churn distribution:\n{df['churned'].value_counts(normalize=True).mul(100).round(2).to_string()}")
print(f"\n[DATA]  Numerical summary:\n{df.describe().T.to_string()}")

# -----------------------------------------------------------------------------
# 3. FEATURE ENGINEERING & PRE-PROCESSING
# -----------------------------------------------------------------------------
df_model = df.drop(columns=["customer_id"]).copy()

# --- Derived features ---
df_model["engagement_ratio"]   = (df_model["avg_watch_time_per_day"] /
                                   (df_model["last_login_days"] + 1))
df_model["is_inactive"]        = (df_model["last_login_days"] > 30).astype(int)
df_model["high_value"]         = (df_model["monthly_fee"] >= 17.99).astype(int)
df_model["multi_profile_user"] = (df_model["number_of_profiles"] > 1).astype(int)

# --- Encode categorical columns ---
cat_cols = ["gender", "subscription_type", "region",
            "device", "payment_method", "favorite_genre"]
le = LabelEncoder()
for col in cat_cols:
    df_model[col] = le.fit_transform(df_model[col].astype(str))

# --- Split features / target ---
X = df_model.drop(columns=["churned"])
y = df_model["churned"]

FEATURE_NAMES = list(X.columns)

# --- Train / test split (stratified 80/20) ---
X_train, X_test, y_train, y_test = train_test_split(
    X, y, test_size=0.20, random_state=42, stratify=y
)

# --- Scaling (for Logistic Regression pipeline) ---
scaler = StandardScaler()
X_train_sc = scaler.fit_transform(X_train)
X_test_sc  = scaler.transform(X_test)

print(f"\n[SPLIT] Train size : {X_train.shape[0]}  |  Test size : {X_test.shape[0]}")

# -----------------------------------------------------------------------------
# 4. MODEL TRAINING
# -----------------------------------------------------------------------------

# -- 4a. Random Forest (primary model) ----------------------------------------
rf = RandomForestClassifier(
    n_estimators   = 200,
    max_depth      = 12,
    min_samples_leaf = 4,
    class_weight   = "balanced",
    random_state   = 42,
    n_jobs         = -1
)
rf.fit(X_train, y_train)
y_pred_rf   = rf.predict(X_test)
y_prob_rf   = rf.predict_proba(X_test)[:, 1]

# -- 4b. Logistic Regression (baseline) ---------------------------------------
lr = LogisticRegression(max_iter=1000, class_weight="balanced", random_state=42)
lr.fit(X_train_sc, y_train)
y_pred_lr   = lr.predict(X_test_sc)
y_prob_lr   = lr.predict_proba(X_test_sc)[:, 1]

# -- 4c. 5-fold cross-validation on RF ----------------------------------------
cv = StratifiedKFold(n_splits=5, shuffle=True, random_state=42)
cv_scores = cross_val_score(rf, X, y, cv=cv, scoring="accuracy")

print(f"\n[CV]    Random Forest 5-fold CV accuracy: "
      f"{cv_scores.mean():.4f} +/- {cv_scores.std():.4f}")

# -----------------------------------------------------------------------------
# 5. MODEL EVALUATION
# -----------------------------------------------------------------------------

def evaluate(name, y_true, y_pred, y_prob):
    acc  = accuracy_score(y_true, y_pred)
    prec = precision_score(y_true, y_pred)
    rec  = recall_score(y_true, y_pred)
    f1   = f1_score(y_true, y_pred)
    auc  = roc_auc_score(y_true, y_prob)
    cm   = confusion_matrix(y_true, y_pred)
    return dict(name=name, acc=acc, prec=prec, rec=rec, f1=f1, auc=auc, cm=cm)

rf_metrics = evaluate("Random Forest",       y_test, y_pred_rf, y_prob_rf)
lr_metrics = evaluate("Logistic Regression", y_test, y_pred_lr, y_prob_lr)

for m in [rf_metrics, lr_metrics]:
    print(f"\n{'-'*55}")
    print(f"  Model      : {m['name']}")
    print(f"  Accuracy   : {m['acc']:.4f}  ({m['acc']*100:.2f}%)")
    print(f"  Precision  : {m['prec']:.4f}")
    print(f"  Recall     : {m['rec']:.4f}")
    print(f"  F1-Score   : {m['f1']:.4f}")
    print(f"  ROC-AUC    : {m['auc']:.4f}")
    print(f"  Confusion Matrix:\n{m['cm']}")

print(f"\n[REPORT] Full classification report - Random Forest:\n")
print(classification_report(y_test, y_pred_rf, target_names=["Retained", "Churned"]))

# Feature importance (RF)
feat_imp = pd.Series(rf.feature_importances_, index=FEATURE_NAMES).sort_values(ascending=False)
print("\n[FEATURES] Top-10 feature importances (Random Forest):")
print(feat_imp.head(10).to_string())

# ROC curve data
fpr_rf, tpr_rf, _ = roc_curve(y_test, y_prob_rf)
fpr_lr, tpr_lr, _ = roc_curve(y_test, y_prob_lr)

# -----------------------------------------------------------------------------
# 6. EXECUTIVE 4-PANEL DASHBOARD
# -----------------------------------------------------------------------------
fig, axes = plt.subplots(2, 2, figsize=(16, 12))
fig.patch.set_facecolor(NETFLIX_DARK)
fig.suptitle(
    "Netflix Customer Churn -- Executive Dashboard",
    fontsize=20, fontweight="bold", color="white", y=0.98
)

for ax in axes.flat:
    ax.set_facecolor("#1C1C1C")
    for spine in ax.spines.values():
        spine.set_edgecolor("#444444")

# -- Panel 1: Confusion Matrix (Random Forest) --------------------------------
ax1 = axes[0, 0]
cm_rf = rf_metrics["cm"]
sns.heatmap(
    cm_rf,
    annot=True, fmt="d", cmap="Reds",
    xticklabels=["Retained", "Churned"],
    yticklabels=["Retained", "Churned"],
    linewidths=0.5, linecolor="#333",
    ax=ax1, cbar=False, annot_kws={"size": 14, "weight": "bold"}
)
ax1.set_title("Confusion Matrix -- Random Forest", color="white", fontsize=13, pad=10)
ax1.set_xlabel("Predicted",  color="#CCCCCC", fontsize=11)
ax1.set_ylabel("Actual",     color="#CCCCCC", fontsize=11)
ax1.tick_params(colors="#CCCCCC")

# annotate TN/FP/FN/TP
labels = [["TN", "FP"], ["FN", "TP"]]
for i in range(2):
    for j in range(2):
        ax1.text(j + 0.5, i + 0.75, labels[i][j],
                 ha="center", va="center", fontsize=9,
                 color="white", alpha=0.6)

# -- Panel 2: ROC Curves -------------------------------------------------------
ax2 = axes[0, 1]
ax2.plot(fpr_rf, tpr_rf, color=NETFLIX_RED,  lw=2.2,
         label=f"Random Forest  (AUC = {rf_metrics['auc']:.3f})")
ax2.plot(fpr_lr, tpr_lr, color="#F5A623",    lw=2.2, linestyle="--",
         label=f"Logistic Reg.  (AUC = {lr_metrics['auc']:.3f})")
ax2.plot([0, 1], [0, 1], color="#666666", lw=1.2, linestyle=":")
ax2.fill_between(fpr_rf, tpr_rf, alpha=0.08, color=NETFLIX_RED)
ax2.set_xlim([0, 1]); ax2.set_ylim([0, 1.02])
ax2.set_title("ROC Curve Comparison", color="white", fontsize=13, pad=10)
ax2.set_xlabel("False Positive Rate", color="#CCCCCC", fontsize=11)
ax2.set_ylabel("True Positive Rate",  color="#CCCCCC", fontsize=11)
ax2.tick_params(colors="#CCCCCC")
ax2.legend(facecolor="#2A2A2A", labelcolor="white", fontsize=10, loc="lower right")

# -- Panel 3: Feature Importance (top 10) -------------------------------------
ax3 = axes[1, 0]
top10 = feat_imp.head(10)
colors_bar = [NETFLIX_RED if i == 0 else "#9B0000" if i < 3 else NETFLIX_GREY
              for i in range(len(top10))]
bars = ax3.barh(top10.index[::-1], top10.values[::-1],
                color=colors_bar[::-1], edgecolor="#333", height=0.7)
for bar, val in zip(bars, top10.values[::-1]):
    ax3.text(val + 0.001, bar.get_y() + bar.get_height() / 2,
             f"{val:.3f}", va="center", ha="left",
             color="white", fontsize=9)
ax3.set_title("Top-10 Feature Importances (Random Forest)",
              color="white", fontsize=13, pad=10)
ax3.set_xlabel("Importance Score", color="#CCCCCC", fontsize=11)
ax3.tick_params(colors="#CCCCCC")
ax3.set_xlim(0, top10.values.max() * 1.18)

# -- Panel 4: Model Metrics Comparison (grouped bar) --------------------------
ax4 = axes[1, 1]
metrics_labels  = ["Accuracy", "Precision", "Recall", "F1-Score", "ROC-AUC"]
rf_vals = [rf_metrics["acc"], rf_metrics["prec"], rf_metrics["rec"],
           rf_metrics["f1"],  rf_metrics["auc"]]
lr_vals = [lr_metrics["acc"], lr_metrics["prec"], lr_metrics["rec"],
           lr_metrics["f1"],  lr_metrics["auc"]]

x_pos  = np.arange(len(metrics_labels))
width  = 0.35
b1 = ax4.bar(x_pos - width/2, rf_vals, width,
             label="Random Forest",       color=NETFLIX_RED,  alpha=0.9, edgecolor="#333")
b2 = ax4.bar(x_pos + width/2, lr_vals, width,
             label="Logistic Regression", color="#F5A623", alpha=0.9, edgecolor="#333")

for bar in list(b1) + list(b2):
    ax4.text(bar.get_x() + bar.get_width() / 2,
             bar.get_height() + 0.005,
             f"{bar.get_height():.2f}",
             ha="center", va="bottom", color="white", fontsize=8.5)

ax4.set_ylim(0, 1.12)
ax4.set_xticks(x_pos)
ax4.set_xticklabels(metrics_labels, color="#CCCCCC", fontsize=10)
ax4.set_title("Model Performance Comparison", color="white", fontsize=13, pad=10)
ax4.set_ylabel("Score", color="#CCCCCC", fontsize=11)
ax4.tick_params(colors="#CCCCCC")
ax4.legend(facecolor="#2A2A2A", labelcolor="white", fontsize=10)
ax4.axhline(y=0.80, color="#FFFFFF", linestyle="--", lw=0.8, alpha=0.3)

# -- Annotation bar at bottom -------------------------------------------------
fig.text(
    0.5, 0.01,
    f"Dataset: netflix_customer_churn.csv  |  Records: 5,000  |  "
    f"RF 5-Fold CV: {cv_scores.mean():.4f} +/- {cv_scores.std():.4f}  |  "
    f"Train/Test split: 80/20  |  Author: Sahith Uppala",
    ha="center", fontsize=9, color="#888888"
)

plt.tight_layout(rect=[0, 0.03, 1, 0.96])
plt.savefig(DASHBOARD_PATH, dpi=180, bbox_inches="tight",
            facecolor=NETFLIX_DARK)
plt.close()
print(f"\n[DASHBOARD]  Saved -> {DASHBOARD_PATH}")

# ---------------------------------------------------------------------------
# 7. CONSOLE REPORT SUMMARY
# ---------------------------------------------------------------------------
print("\n" + "=" * 70)
print("  FINAL EXECUTIVE SUMMARY")
print("=" * 70)
print(f"\n  Dataset         : netflix_customer_churn.csv")
print(f"  Total records   : {len(df):,}")
print(f"  Churn rate      : {df['churned'].mean()*100:.1f}%")
print(f"  Features used   : {len(FEATURE_NAMES)}  (incl. 4 engineered features)")
print(f"  Train/Test split: 80% / 20%  (stratified)")
print()
print(f"  +-------------------------+----------+----------+")
print(f"  | Metric                  | Rand. F. | Log. Reg |")
print(f"  +-------------------------+----------+----------+")
for label, k in [("Accuracy ", "acc"), ("Precision", "prec"),
                 ("Recall   ", "rec"), ("F1-Score ", "f1"), ("ROC-AUC  ", "auc")]:
    print(f"  | {label}                 |  {rf_metrics[k]:.4f}  |  {lr_metrics[k]:.4f}  |")
print(f"  +-------------------------+----------+----------+")
print()
print(f"  5-Fold CV (Random Forest): {cv_scores.mean():.4f} +/- {cv_scores.std():.4f}")
print(f"\n  Dashboard saved : {DASHBOARD_PATH}")
print("\n" + "=" * 70)
print("  Run complete. All outputs ready for report submission.")
print("=" * 70 + "\n")
