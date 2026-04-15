# Presentation Script — Predicting Car Crash Severity
**Target time: 10 minutes | ~1,400 words**

---

## SLIDE 1 — Title  *(0:00 – 0:20)*

> "Good [morning/afternoon]. Our project is called *Predicting Car Crash Severity:
> A Data Mining Analysis of California Traffic Incidents.*
> I'm [Name], and I'm joined by Keaton, Tyler, Sahil, Will, and Alex.
> We'll walk you through our full CRISP-DM analysis in about ten minutes."

---

## SLIDE 2 — Agenda  *(0:20 – 0:35)*

> "Here's our roadmap. We'll cover the business problem, the dataset, how we prepared
> the data, the five models we built, our results, and four concrete recommendations.
> Let's get into it."

---

## SLIDE 3 — Business Understanding  *(0:35 – 1:15)*

> "California sees thousands of traffic fatalities every year. The costs go beyond
> the individuals involved — emergency services get strained, insurance rates climb,
> and infrastructure wears down.
>
> Our analysis is driven by three questions:
> First, what factors actually predict whether a crash is severe or minor?
> Second, are there natural risk profiles we can target?
> And third, how reliably can we predict severity with the data we have?
>
> The answers matter to transportation agencies, insurers, dispatchers, and city planners."

---

## SLIDE 4 — Data Understanding  *(1:15 – 2:00)*

> "Our dataset contains 112,660 California crash records across 11 variables —
> county, day of week, violation category, weather, crash type, whether it was on a
> highway, and whether it occurred in daylight.
>
> The target variable is binary: zero for minor, one for severe.
>
> But here's the catch — and it shapes every decision we made downstream."

---

## SLIDE 5 — Class Imbalance  *(2:00 – 2:45)*

> "93% of crashes are minor. Only 7% — about 8,200 records — are severe.
>
> That means a model that predicts *every single crash as minor* would hit 93% accuracy
> and be completely useless for our actual goal.
>
> So we threw out accuracy as our primary metric and focused on AUC-ROC, F1-score,
> and recall for the severe class. We also applied techniques like minority upsampling
> and balanced class priors to push the models to actually learn the rare severe cases."

---

## SLIDE 6 — Data Preparation  *(2:45 – 3:20)*

> "Our prep pipeline ran in three stages.
>
> First, we dropped the ID column and city — city had hundreds of unique values that
> would inflate the feature space without adding anything county doesn't already capture.
>
> Second, we one-hot encoded CrashType and ViolCat, giving us a 25-column intermediate
> dataset suitable for tree-based models.
>
> Third, for linear and distance-based models, we fully encoded County, Weekday, and Month
> to avoid implying false ordinal relationships.
>
> The result: 112,660 records and 98 fully binary features fed into every model."

---

## SLIDE 7 — Modeling Overview  *(3:20 – 3:50)*

> "We ran five models total — four supervised classifiers to predict severity,
> plus K-Means clustering to find natural groupings in the data without using the
> severity label at all.
>
> Each model had a different strategy for handling the imbalance problem,
> and we'll see how that played out in the results."

---

## SLIDE 8 — Decision Tree  *(3:50 – 5:00)*

> "The Decision Tree is our best model.
>
> We pruned it to a max depth of 6 to prevent overfitting, upsampled the minority class,
> and validated with 5-fold cross-validation.
>
> The results: 89% accuracy, F1 of 0.853, and an AUC-ROC of 0.838.
> Cross-validated F1 held steady at 0.812 — so this isn't a fluke.
>
> Looking at feature importances, **Highway** is the single strongest predictor at 0.142,
> followed by **Daylight** at 0.119 and **ClearWeather** at 0.089.
>
> This makes intuitive sense — highway crashes at night in adverse weather are
> a fundamentally different risk environment.
>
> The tree structure is also fully interpretable, which matters if you want to deploy
> this in a real dispatch system and have stakeholders trust it."

---

## SLIDE 9 — Logistic Regression & Naive Bayes  *(5:00 – 6:00)*

> "Our next two models — Logistic Regression and Naive Bayes — both land in
> 'acceptable' territory.
>
> Logistic Regression hit an AUC of 0.753 and caught 66% of severe crashes.
> It's particularly useful because the model's coefficients directly tell you
> which features contribute most to severity risk — that's actionable for policy.
>
> Naive Bayes achieved an AUC of 0.745 and 64% recall.
> We used the Bernoulli variant because all 98 features are binary — a natural fit.
> No scaling required, and it converges instantly on a dataset this size.
>
> Both models are suitable for probability-based risk scoring, even if neither
> matches the Decision Tree."

---

## SLIDE 10 — KNN & K-Means  *(6:00 – 7:00)*

> "KNN is a cautionary tale.
>
> On the surface it looks great — 92% accuracy. But it identified only 142 of 2,375
> severe crashes in the test set. That's a 6% recall. 94% of severe crashes would be
> missed entirely.
>
> The reason is the imbalance. In any neighborhood of five points drawn from a 93/7
> dataset, minor crashes dominate nearly every vote.
> A McNemar test confirmed the difference from Logistic Regression is
> statistically significant — p of 2 times 10 to the negative 184.
> KNN is not recommended without SMOTE or threshold correction.
>
> On the unsupervised side, K-Means found three meaningful crash archetypes.
> Cluster 1 — highway crashes with high violation rates — has the highest severity
> proportion at 8.9% and is the clear high-priority segment.
> Cluster 0 is your lower-risk non-highway daytime crashes at 5.2%.
> These profiles aren't just academic — they tell you where to direct resources."

---

## SLIDE 11 — Model Comparison  *(7:00 – 7:30)*

> "Here's the full picture side by side.
>
> Decision Tree leads at AUC 0.838. Logistic Regression and Naive Bayes are
> comparable at 0.753 and 0.745. KNN at 0.607 is barely above random chance
> on what actually matters.
>
> The Decision Tree is our recommendation for deployment."

---

## SLIDE 12 — Key Findings  *(7:30 – 8:00)*

> "Circling back to our three business questions:
>
> What predicts severity? Highway status, lighting, weather, and specific
> violation categories — consistent across both the tree and regression coefficients.
>
> Are there distinct profiles? Yes — three clusters with meaningfully different
> risk signatures, ready for targeted intervention.
>
> How reliable is the prediction? An AUC of 0.838 is solid for a real-world,
> heavily imbalanced problem. It far exceeds chance and is sufficient for triage."

---

## SLIDE 13 — Recommendations  *(8:00 – 9:00)*

> "Four recommendations:
>
> **One** — deploy the Decision Tree as a real-time severity triage tool in emergency
> dispatch. It can flag likely-severe crashes the moment an incident is reported.
>
> **Two** — use Logistic Regression coefficients to guide where enforcement
> campaigns and infrastructure dollars go.
>
> **Three** — design cluster-specific prevention programs. Don't apply a uniform
> intervention to all crashes. Cluster 1 highway incidents need a different
> response than Cluster 0 surface-street incidents.
>
> **Four** — address the class imbalance in future iterations.
> SMOTE, threshold calibration, and ensemble methods like XGBoost could push
> severe-class recall meaningfully higher."

---

## SLIDE 14 — Limitations & Future Work  *(9:00 – 9:25)*

> "A few honest limitations.
>
> We don't have continuous features like speed or driver age — those would likely
> improve the model significantly. The binary encoding loses within-county
> geographic nuance. And all classifiers still miss 12–35% of severe crashes
> due to the imbalance ceiling.
>
> Future work: SMOTE, gradient boosting, SHAP explainability,
> and a geospatial clustering layer to replace county codes."

---

## SLIDE 15 — Q&A  *(9:25 – 10:00)*

> "To wrap up — we analyzed 112,660 crash records, built five models,
> and showed that crash severity is genuinely predictable.
>
> The Decision Tree achieves an AUC of 0.838 with an interpretable structure
> ready for real-world deployment. Highway status, nighttime conditions,
> and violation categories are the dominant risk signals.
>
> Thank you — we're happy to take any questions."

---

## Timing Summary

| Slide | Topic | Time | Duration |
|-------|-------|------|----------|
| 1 | Title | 0:00 | 20 sec |
| 2 | Agenda | 0:20 | 15 sec |
| 3 | Business Understanding | 0:35 | 40 sec |
| 4 | Data Understanding | 1:15 | 45 sec |
| 5 | Class Imbalance | 2:00 | 45 sec |
| 6 | Data Preparation | 2:45 | 35 sec |
| 7 | Modeling Overview | 3:20 | 30 sec |
| 8 | Decision Tree | 3:50 | 70 sec |
| 9 | Logistic Regression & NB | 5:00 | 60 sec |
| 10 | KNN & K-Means | 6:00 | 60 sec |
| 11 | Model Comparison | 7:00 | 30 sec |
| 12 | Key Findings | 7:30 | 30 sec |
| 13 | Recommendations | 8:00 | 60 sec |
| 14 | Limitations | 9:00 | 25 sec |
| 15 | Q&A | 9:25 | 35 sec |

**Tips:**
- Slides 8–10 are the technical core — don't rush them.
- If you're running long, compress slides 2, 7, and 14 first.
- Have one person own the clicker and one person ready to answer the KNN/imbalance question — it always comes up.
