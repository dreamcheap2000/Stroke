# Dataset Specification & GitHub Copilot Guide: 20260806_DeID.csv

This specification document introduces the post-stroke clinical registry dataset `20260806_DeID.csv` to GitHub Copilot (or other AI coding assistants). It maps the study cohort context, chronological timelines, variable schemas, and de-identification characteristics to facilitate seamless data processing, statistical modeling, and machine learning.

---

## 1. Study Cohort Context & Timeline

*   **Clinical Registry**: This dataset represents a prospective de-identified cohort of **633 consecutive stroke patients** undergoing intensive rehabilitation under the Post-Acute Care Cerebrovascular Diseases (PAC-CVD) program at Taoyuan Chang Gung Memorial Hospital, following hyperacute stabilization at Linkou Chang Gung Memorial Hospital (enrolled between 2014 and 2019).
*   **Program Completion and Attrition**:
    *   **Completers ($N = 577, 91.2\%$)**: Patients who successfully completed the longitudinal intensive rehabilitation protocol.
    *   **Non-Completers ($N = 56, 8.8\%$)**: Patients who withdrew prematurely due to acute medical deterioration ($n = 31$), discharge against medical advice ($n = 18$), or death ($n = 7$).
*   **Chronological Framework**:
    *   **T0 (Acute Ward Discharge / Day of Transfer)**: Deficit severity is assessed using individual NIHSS sub-items immediately prior to discharge from the acute neurology ward.
    *   **T1 (PAC Admission / Day 1)**: Patients transfer to the post-acute rehabilitation ward. Comprehensive baseline demographics, comorbidities, functional status, balance, and nutritional profiles are evaluated before intensive therapies begin.
    *   **T2 (Week 3), T3 (Week 6), T4 (Discharge / Up to 12 Weeks)**: Specially timed follow-up checkpoints tracking physical and functional recovery.

---

## 2. Clinical Variable Categories & Schema Mappings

The registry columns are grouped into nine functional domains:

### Category A: Demographics & Core Metadata
*   `ID`: De-identified unique patient registry identifier.
*   `Age`: Patient age (continuous, in years).
*   `Sex, F0 M1`: Biological sex (binary: $0 = \text{Female}$, $1 = \text{Male}$).
*   `RehabDate` & `StrokeDate`: Dates of rehabilitation admission and index stroke onset.

### Category B: Stroke Etiology & Anatomical Topology
*   `HemorrhageStroke`: Biological stroke mechanism (binary: $1 = \text{Hemorrhagic}$, $0 = \text{Ischemic}$).
*   `SideOfStroke1Rt2Lt3Bil` (or `Side_Right`/`Side_Left`): Lesion hemispheric side (categorical: $1 = \text{Right}$, $2 = \text{Left}$, $3 = \text{Bilateral}$).
*   `Stroke Location` (`Loc_CortSub`, `Loc_Subcortical`, `Loc_Infratentorial`): Brain lesion region (categorical).
*   `Dissection`: Arterial dissection etiology (binary).
*   `ACA`: Involvement of the Anterior Cerebral Artery vascular territory (binary).
*   `LVS` & `LVO`: Large Vessel Stenosis and Large Vessel Occlusion markers (binary).
*   `Undetermined`: Ischemic stroke of undetermined etiology (binary).

### Category C: Baseline Comorbidities (16 Binary Variables)
All comorbidities are recorded at PAC admission (T1) as binary flags ($1 = \text{Present}$, $0 = \text{Absent}$):
*   `AF` (Atrial Fibrillation), `DM` (Diabetes Mellitus), `HTN` (Hypertension), `Dyslipidemia`, `CAD` (Coronary Artery Disease), `CKD` (Chronic Kidney Disease), `RestrictiveLung` (Restrictive Lung Disease), `GIUlcer` (Gastroduodenal Ulcer), `LiverCirrhosis`, `Hepatitis`, `Parkinsonism`, `Malignancy`, `OldStroke` (History of Prior Stroke), `Dementia`, `Psychiatric` (including bipolar, delusion, and anxiety disorders), and `Gout`.

### Category D: Acute Ward Management & Complications (T0 Pre-Transfer)
Complications and treatments arising on the acute ward prior to PAC transfer:
*   `Pneumonia`, `UTI` (Urinary Tract Infection), `GIB` (Gastrointestinal Bleeding), `Cellulitis` (all binary).
*   `tPA`, `IA`, `tPAIA`: Hyperacute therapies (Intravenous tPA, Intra-arterial Thrombectomy, or Combined therapy).
*   `Neurology_LOS`: Acute neurology ward length of stay (continuous, in days).

### Category E: Acute Discharge Neurological Deficits (T0 NIHSS Sub-items)
Granular neurological deficits evaluated at acute-ward discharge to capture focal impairment:
*   `ConsOut`: Level of consciousness responsiveness ($0$ to $3$).
*   `AnswerOut` & `OrderOut`: LOC questions and commands ($0$ to $2$).
*   `EOMOut`: Best gaze / extraocular movement abnormalities ($0$ to $2$).
*   `VisualOut`: Visual fields ($0$ to $3$).
*   `FacialOut`: Facial palsy ($0$ to $3$).
*   `LUOut` & `RUOut`: Left and Right upper extremity motor drift ($0$ to $4$).
*   `LLOut` & `RLOut`: Left and Right lower extremity motor drift ($0$ to $4$).
*   `Coordinateout`: Limb ataxia and coordination ($0$ to $2$).
*   `SensoryOut`: Somatosensation loss ($0$ to $2$).
*   `LanguageOut`: Best language and aphasia severity ($0$ to $3$).
*   `ArticulateOut`: Dysarthria ($0$ to $2$).
*   `NeglectOut`: Extinction, inattention, or hemispatial neglect ($0$ to $2$).

### Category F: PAC Ward Functional Assessments (T1 Admission Baseline)
Standardized functional, balance, cognitive, and nutritional scales assessed on Day 1 of transfer:
*   `BBS1`: Berg Balance Scale score ($0$ to $56$; higher indicates better static/dynamic balance).
*   `BI1`: Barthel Index of Activities of Daily Living ($0$ to $100$; higher indicates greater self-care independence).
*   `MNA1`: Mini-Nutritional Assessment ($0$ to $30$; $<17 = \text{malnutrition}$, $17\text{--}23 = \text{at risk}$, $\ge24 = \text{normal}$).
*   `EuroQoL5D1`: EuroQol-5D Quality of Life Index ($5$ to $15$; higher indicates poorer quality of life).
*   `FuglUE1` & `FuglSEN1`: Fugl-Meyer Upper Extremity Motor score ($0$ to $66$) and Sensory score ($0$ to $24$).
*   `IADL1`: Instrumental Activities of Daily Living ($0$ to $8$).
*   `FOIS1`: Functional Oral Intake Scale for dysphagia ($1$ to $7$).
*   `MRS1`: modified Rankin Scale of global disability ($0$ to $6$).
*   `CCAT1`: Concise Chinese Aphasia Test score ($0$ to $12$).
*   `MMSE1` & `MAL1`: Mini-Mental State Examination and Motor Activity Log (subject to administrative missingness).

### Category G: Longitudinal Outcomes & Locomotor Milestones (T1 to T4)
Longitudinal measures tracking physical locomotion progress over 12 weeks:
*   `6MWT1`, `6MWT2`, `6MWT3`, `6MWT4`: 6-Minute Walk Test distance in meters, measured at admission (T1), Week 3 (T2), Week 6 (T3), and discharge (T4).
*   `Gait_Speed_1`, `Gait_Speed_2`, `Gait_Speed_3`, `Gait_Speed_4`: Comfortable gait speed (m/s) at corresponding intervals.
*   `First_6MWT_TP`: The specific rehabilitation milestone timepoint when independent walking (6MWT > 0) was first achieved (`T1`, `T2`, `T3`, `T4`, or `Never`).

### Category H: T1-to-T2 Dynamic Functional Improvements (Rehabilitation Deltas)
These variables capture early therapeutic responsiveness and are calculated as the longitudinal score at Week 3 (T2) minus the baseline score at PAC Admission (T1) [i.e., T2 - T1]. These continuous deltas measure early physiological recovery:
1.  **Locomotor & Postural Deltas**:
    *   `6MWT_Improvement_T1T2` (or `6MWT_Imp_T1T2`): Early 3-Week Walking Endurance Improvement (meters).
    *   `BBS_Improvement_T1T2`: Early 3-Week Postural Control & Balance Progress (Berg Balance Scale points).
2.  **Activities of Daily Living & Disability Deltas**:
    *   `BI_Improvement_T1T2` (or `BI_Change_T1T2`): Early 3-Week Daily Life Independence Progress (Barthel Index points).
    *   `MRS_Improvement_T1T2` (or `MRS_Change_T1T2`): Early 3-Week Global Disability Reduction (modified Rankin Scale points change).
3.  **Nutritional, Motor & Sensory Deltas**:
    *   `MNA_Improvement_T1T2` (or `MNA_Change_T1T2`): Early 3-Week Nutritional Status Improvement (Mini-Nutritional Assessment points).
    *   `FuglUE_Improvement_T1T2`: Early 3-Week Paretic Upper Limb Synergism/Motor Progress (Fugl-Meyer Upper Extremity points).
    *   `FuglSEN_Improvement_T1T2`: Early 3-Week Somatosensory Recovery Progress (modified Fugl-Meyer Sensory scale points).
    *   `FOIS_Improvement_T1T2`: Early 3-Week Swallowing/Oral Intake Recovery Progress (Functional Oral Intake Scale points).

### Category I: Rehabilitation Metrics & Metadata
*   `Rehab_LOS` & `Rehab_LOS_Category`: Total PAC length of stay (days) and category (Short: $\le21$ days, Intermediate: $22\text{--}42$ days, Long: $>42$ days).
*   `PAC_Program_Completion`: Final administrative status (`Completed PAC program` vs. `Did not complete PAC program`).
*   `DischargeDestination`: Categorical tracking of discharge placement ($1 = \text{home}$, $2 = \text{hospital}$, $3 = \text{nursing facility}$, $4 = \text{rehab}$, $5 = \text{discharge against medical advice}$, $6 = \text{Chinese medicine ward}$, $7 = \text{death}$).

---

## 3. Structural Anomalies & EHR-Specific Characteristics

To prevent processing bias, the following data-cleaning properties of `20260806_DeID.csv` must be acknowledged:

1.  **Physical "Structural Zeroes" in Gait Speed & 6MWT**:
    Patients who are completely non-ambulatory at admission cannot physically stand or walk for clinical tests. Standard medical registry practice records these as `NaN` or missing. In this dataset, a blank or `Never` value in walking milestones means the patient's comfortable gait speed (`Gait_Speed_1`) and 6MWT distance (`6MWT1`) should be structurally treated as `0.0` (not missing or dropped) to avoid bias toward high-functioning patients.
2.  **Administrative Policy Missingness (MMSE1 & MAL1)**:
    Due to an institutional EHR policy update on January 1, 2018, the `MMSE` and `MAL` scales were retired from standard post-stroke logging protocols. This introduced exactly **39.3% random missingness** for these variables among program completers. Analyses seeking to maximize sample size (e.g., maintaining N=511 completers) traditionally exclude these features to avoid massive listwise deletion.
