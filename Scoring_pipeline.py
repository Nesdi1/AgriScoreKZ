
import pandas as pd
import numpy as np
import warnings, os, json
from datetime import date
warnings.filterwarnings('ignore')

MORTALITY_NORMS = {
    'скотоводство':    2.0,
    'овцеводство':     3.0,
    'козоводство':     3.0,
    'коневодство':     2.3,
    'верблюдоводство': 6.0,
    'птицеводство':    7.5,
    'свиноводство':    5.2,
    'пчеловодство':   20.0,
}

PASTURE_NORMS_HA = {
    'Костанайская область':           6.0,
    'Северо-Казахстанская область':   5.0,
    'Акмолинская область':            6.5,
    'Павлодарская область':           9.5,
    'Карагандинская область':        12.0,
    'область Абай':                   9.0,
    'Восточно-Казахстанская область': 7.5,
    'Актюбинская область':           13.0,
    'Западно-Казахстанская область': 10.0,
    'Атырауская область':            18.0,
    'Мангистауская область':         12.0,
    'Алматинская область':           10.5,
    'область Жетісу':                10.5,
    'Жамбылская область':            11.0,
    'Туркестанская область':         10.0,
    'Кызылординская область':        14.0,
    'область Ұлытау':                13.5,
    'г.Шымкент':                      5.0,
    'г.Астана':                       5.0,
    'г.Алматы':                       5.0,
}
DEFAULT_PASTURE_NORM = 10.0

DIRECTION_TO_ANIMAL = {
    'Субсидирование в скотоводстве':                          'скотоводство',
    'Субсидирование в овцеводстве':                           'овцеводство',
    'Субсидирование в козоводстве':                           'козоводство',
    'Субсидирование в коневодстве':                           'коневодство',
    'Субсидирование в верблюдоводстве':                       'верблюдоводство',
    'Субсидирование в птицеводстве':                          'птицеводство',
    'Субсидирование в свиноводстве':                          'свиноводство',
    'Субсидирование в пчеловодстве':                          'пчеловодство',
    'Субсидирование затрат по искусственному осеменению':     'скотоводство',
}

DIRECTION_PRIORITY = {
    'Субсидирование в скотоводстве':                          3,
    'Субсидирование затрат по искусственному осеменению':     3,
    'Субсидирование в овцеводстве':                           2,
    'Субсидирование в коневодстве':                           2,
    'Субсидирование в птицеводстве':                          2,
    'Субсидирование в верблюдоводстве':                       2,
    'Субсидирование в козоводстве':                           1,
    'Субсидирование в свиноводстве':                          1,
    'Субсидирование в пчеловодстве':                          1,
}

REJECT_STATUSES = ['Отозвано', 'Отклонена']

print("=" * 65)
print("  ФИНАЛЬНЫЙ СКОРИНГ — Merit-Based + НПА")
print("=" * 65)

print("\n▶ ШАГ 1: Загрузка данных")

possible_names = [
    "Выгрузка_по_выданным_субсидиям_2025_год__обезлич_.xlsx",
    "Выгрузка по выданным субсидиям 2025 год (обезлич).xlsx",
    "/mnt/user-data/uploads/Выгрузка_по_выданным_субсидиям_2025_год__обезлич_.xlsx",
]

df = None
for name in possible_names:
    if os.path.exists(name):
        print(f"  Файл: {name}")
        df = pd.read_excel(name, header=4, engine='openpyxl')
        break

if df is None:
    raise FileNotFoundError("Excel файл не найден — укажите путь вручную")

df.columns = [
    'id', 'date', 'col2', 'col3', 'region', 'akimat',
    'app_num', 'direction', 'subsidy_name', 'status',
    'normative', 'amount', 'district'
]
df = df.drop(columns=['col2', 'col3'])
df = df.dropna(subset=['id'])
df['date']      = pd.to_datetime(df['date'],      errors='coerce')
df['amount']    = pd.to_numeric(df['amount'],     errors='coerce').fillna(0)
df['normative'] = pd.to_numeric(df['normative'],  errors='coerce').fillna(0)
df = df.sort_values('date').reset_index(drop=True)

print(f"  Строк: {len(df):,} | Период: {df['date'].min().date()} — {df['date'].max().date()}")

print("\n▶ ШАГ 2: Временной split (80/20)")

split_idx = int(len(df) * 0.80)
train_df  = df.iloc[:split_idx].copy()
test_df   = df.iloc[split_idx:].copy()

for d in [train_df, test_df, df]:
    d['target'] = (d['status'] == 'Исполнена').astype(int)

global_mean = train_df['target'].mean()
print(f"  Train: {len(train_df):,} | Test: {len(test_df):,} | Target mean: {global_mean:.1%}")

print("\n▶ ШАГ 3А: Eligibility Filter (НПА №18404)")

DIRECTION_MIN_NORMATIVE = {
    'Субсидирование в скотоводстве':                      10_000,
    'Субсидирование в овцеводстве':                        5_000,
    'Субсидирование в козоводстве':                        5_000,
    'Субсидирование в коневодстве':                       10_000,
    'Субсидирование в верблюдоводстве':                   10_000,
    'Субсидирование в птицеводстве':                         200,
    'Субсидирование в свиноводстве':                       5_000,
    'Субсидирование в пчеловодстве':                         200,
    'Субсидирование затрат по искусственному осеменению':    200,
}
DIRECTION_MIN_UNITS = {
    'Субсидирование в скотоводстве':                       1,
    'Субсидирование в овцеводстве':                        5,
    'Субсидирование в козоводстве':                        5,
    'Субсидирование в коневодстве':                        1,
    'Субсидирование в верблюдоводстве':                    1,
    'Субсидирование в птицеводстве':                      10,
    'Субсидирование в свиноводстве':                       5,
    'Субсидирование в пчеловодстве':                       5,
    'Субсидирование затрат по искусственному осеменению':  1,
}

def add_eligibility(d):
    d = d.copy()

    d['is_active_application'] = (~d['status'].isin(REJECT_STATUSES)).astype(int)
    d['units_claimed'] = (
        d['amount'] / d['normative'].replace(0, np.nan)
    ).fillna(0).clip(0, 100_000)
    d['has_valid_amount']    = ((d['amount'] > 0) & (d['normative'] > 0)).astype(int)
    d['has_region']          = d['region'].notna().astype(int)
    d['has_direction']       = d['direction'].notna().astype(int)
    d['has_district']        = d['district'].notna().astype(int)
    d['has_known_direction'] = d['direction'].isin(DIRECTION_TO_ANIMAL).astype(int)

    d['min_normative_ok'] = d.apply(
        lambda r: 1 if (
            pd.notna(r['direction']) and
            r['direction'] in DIRECTION_MIN_NORMATIVE and
            r['normative'] >= DIRECTION_MIN_NORMATIVE[r['direction']]
        ) else 0, axis=1
    )
    d['min_units_ok'] = d.apply(
        lambda r: 1 if (
            pd.notna(r['direction']) and
            r['direction'] in DIRECTION_MIN_UNITS and
            r['units_claimed'] >= DIRECTION_MIN_UNITS[r['direction']]
        ) else 0, axis=1
    )
    d['amount_vs_normative_ok'] = (
        d['amount'] <= d['normative'] * d['units_claimed'] * 1.20 + 1
    ).astype(int)
    d['registry_evidence_flag'] = d['has_known_direction']
    d['missing_regulatory_fields_flag'] = (
        (d['has_district'] == 0) |
        (d['has_known_direction'] == 0) |
        (d['min_normative_ok'] == 0)
    ).astype(int)
    d['legal_compliance_score'] = (
        0.20 * d['has_valid_amount'] +
        0.10 * d['has_region'] +
        0.10 * d['has_direction'] +
        0.05 * d['has_district'] +
        0.10 * d['has_known_direction'] +
        0.20 * d['min_normative_ok'] +
        0.20 * d['min_units_ok'] +
        0.05 * d['amount_vs_normative_ok']
    ).clip(0, 1)

    d['eligibility_score_raw'] = (
        d['is_active_application']    * 0.30 +
        d['has_valid_amount']         * 0.20 +
        d['has_region']               * 0.05 +
        d['has_direction']            * 0.05 +
        d['has_district']             * 0.05 +
        d['has_known_direction']      * 0.05 +
        d['legal_compliance_score']   * 0.30  
    )
    d['eligible_flag'] = (
        (d['is_active_application'] == 1) &
        (d['legal_compliance_score'] >= 0.75)
    ).astype(int)
    return d

for d_name, d_ref in [('train_df', train_df), ('test_df', test_df), ('df', df)]:
    locals()[d_name] if False else None

train_df = add_eligibility(train_df)
test_df  = add_eligibility(test_df)
df       = add_eligibility(df)

print(f"  Eligible: {df['eligible_flag'].sum():,} из {len(df):,} ({df['eligible_flag'].mean():.1%})")

print("\n▶ ШАГ 3Б: Risk Features (НПА №12488 + №11064)")


group_stats = (
    train_df.groupby(['region', 'direction'])['amount']
    .agg(['median', 'mean', 'std', 'count']).reset_index()
)
group_stats.columns = ['region','direction',
                        'grp_amount_median','grp_amount_mean',
                        'grp_amount_std','grp_amount_count']

norm_stats = (
    train_df.groupby('normative')['amount']
    .agg(['median', 'mean', 'std']).reset_index()
)
norm_stats.columns = ['normative',
                       'norm_amount_median','norm_amount_mean','norm_amount_std']

def add_risk_features(d):
    d = d.copy()

    d = d.merge(group_stats, on=['region','direction'], how='left')
    d = d.merge(norm_stats,  on='normative', how='left')

    d['amount_vs_group_median'] = (
        d['amount'] / d['grp_amount_median'].replace(0, np.nan)
    ).fillna(1).clip(0, 20)

    d['amount_vs_norm_median'] = (
        d['amount'] / d['norm_amount_median'].replace(0, np.nan)
    ).fillna(1).clip(0, 20)

    d['group_z_amount'] = (
        (d['amount'] - d['grp_amount_mean']) /
        d['grp_amount_std'].replace(0, np.nan)
    ).fillna(0).clip(-5, 5)

    d['amount_anomaly_flag'] = (
        (d['amount_vs_group_median'] > 3.0) |
        (d['amount_vs_norm_median']  > 3.0) |
        (d['group_z_amount'].abs()   > 3.0)
    ).astype(int)

    d['animal_type']        = d['direction'].map(DIRECTION_TO_ANIMAL).fillna('скотоводство')
    d['mortality_norm_pct'] = d['animal_type'].map(MORTALITY_NORMS).fillna(2.0)
    d['pasture_norm_ha']    = d['region'].map(PASTURE_NORMS_HA).fillna(DEFAULT_PASTURE_NORM)
    d['direction_priority'] = d['direction'].map(DIRECTION_PRIORITY).fillna(1)

    d['required_pasture_ha'] = d['units_claimed'] * d['pasture_norm_ha']
    d['pasture_load_log']    = np.log1p(d['required_pasture_ha'])

    d['expected_mortality']  = d['units_claimed'] * d['mortality_norm_pct'] / 100
    d['mortality_risk_log']  = np.log1p(d['expected_mortality'])

    d['pasture_overload_ratio'] = (d['required_pasture_ha'] / 200.0).clip(0, 10)
    d['pasture_overload_flag']  = (d['required_pasture_ha'] > 200).astype(int)

    d['mortality_deviation_ratio'] = (
        d['expected_mortality'] / d['units_claimed'].replace(0, np.nan)
    ).fillna(0).clip(0, 1)
    d['mortality_high_flag'] = (d['mortality_deviation_ratio'] > 0.02).astype(int)

    d['npa_conflict_count'] = (
        d['amount_anomaly_flag'].fillna(0).astype(int) +
        d['pasture_overload_flag'].fillna(0).astype(int) +
        d['mortality_high_flag'].fillna(0).astype(int)
    )

    d['risk_score_rule'] = (
        0.40 * d['amount_anomaly_flag'].fillna(0).astype(float) +
        0.30 * d['pasture_overload_flag'].fillna(0).astype(float) +
        0.30 * d['mortality_high_flag'].fillna(0).astype(float)
    ).clip(0, 1)

    d['risk_bucket'] = np.select(
        [
            d['risk_score_rule'] >= 0.60,
            d['risk_score_rule'] >= 0.25,
        ],
        ['HIGH', 'MEDIUM'],
        default='LOW'
    )

    d['manual_review_flag'] = np.where(
        (d['npa_conflict_count'] >= 1) |
        (d['risk_score_rule'] >= 0.25) |
        (d['missing_regulatory_fields_flag'] == 1),
        1, 0
    )

    return d

train_df = add_risk_features(train_df)
test_df  = add_risk_features(test_df)
df       = add_risk_features(df)

print(f"  Аномальных сумм:       {df['amount_anomaly_flag'].sum():,}")
print(f"  На ручную проверку:    {df['manual_review_flag'].sum():,}")

print("\n▶ ШАГ 4: Merit Features (сглаженный encoding, без очерёдности)")

SMOOTHING = 50 

def smoothed_enc(train, col):
    s = train.groupby(col)['target'].agg(['mean','count']).reset_index()
    s['enc'] = (s['mean'] * s['count'] + global_mean * SMOOTHING) / (s['count'] + SMOOTHING)
    return s[[col, 'enc', 'count']].rename(columns={'enc': f'{col}_enc_s', 'count': f'{col}_cnt'})

region_enc_df    = smoothed_enc(train_df, 'region')
direction_enc_df = smoothed_enc(train_df, 'direction')
normative_enc_df = smoothed_enc(train_df, 'normative')

def add_merit_features(d, fit=None):
    d = d.copy()
    ref = fit if fit is not None else d

    drop = [c for c in d.columns if c in [
        'region_enc_s','region_cnt','direction_enc_s','direction_cnt',
        'normative_enc_s','normative_cnt',
        'dir_amount_med','region_amount_q75',
        'region_approval_rate','dir_approval_rate','normative_approval_rate',
        'region_total_apps','dir_total_apps',
    ]]
    d.drop(columns=drop, errors='ignore', inplace=True)

    d = d.merge(region_enc_df,    on='region',    how='left')
    d = d.merge(direction_enc_df, on='direction', how='left')
    d = d.merge(normative_enc_df, on='normative', how='left')

    for col in ['region_enc_s','direction_enc_s','normative_enc_s']:
        d[col] = d[col].fillna(global_mean)
    for col in ['region_cnt','direction_cnt','normative_cnt']:
        d[col] = d[col].fillna(0)

    dir_med = ref.groupby('direction')['amount'].median().reset_index()
    dir_med.columns = ['direction','dir_amount_med']
    d = d.merge(dir_med, on='direction', how='left')
    d['amount_vs_dir_med'] = np.log1p(
        d['amount'] / d['dir_amount_med'].clip(lower=1)
    ).fillna(0)

    q75 = ref.groupby('region')['amount'].quantile(0.75).reset_index()
    q75.columns = ['region','region_amount_q75']
    d = d.merge(q75, on='region', how='left')
    d['is_large_farm'] = (d['amount'] >= d['region_amount_q75'].fillna(0)).astype(int)

    d['merit_volume'] = np.log1p(d['amount']) * (1 - d['risk_score_rule'].clip(0, 1))

    d['quarter']      = d['date'].dt.quarter
    d['week_of_year'] = d['date'].dt.isocalendar().week.astype(int)

    d['log_amount']   = np.log1p(d['amount'])

    d['subsidy_per_unit'] = (
        d['amount'] / d['units_claimed'].replace(0, np.nan)
    ).fillna(0).clip(0, 1_000_000)

    d['normative_efficiency'] = (
        1 - (
            (d['subsidy_per_unit'] - d['normative'].replace(0, np.nan)).abs() /
            d['normative'].replace(0, np.nan)
        ).fillna(0).clip(0, 1)
    )

    d['policy_impact_score'] = (
        0.40 * (d['direction_priority'] / 3.0) +
        0.35 * d['normative_efficiency'] +
        0.25 * d['is_large_farm']
    ).clip(0, 1)


    d['budget_efficiency_score'] = (
        0.60 * d['policy_impact_score'] +
        0.40 * d['normative_efficiency']
    ).clip(0, 1)

    return d

train_df = add_merit_features(train_df, fit=train_df)
test_df  = add_merit_features(test_df,  fit=train_df)
df       = add_merit_features(df,       fit=train_df)

from sklearn.preprocessing import LabelEncoder

cat_cols = ['region', 'direction', 'district']
le_dict  = {}
for col in cat_cols:
    le = LabelEncoder()
    le.fit(train_df[col].fillna('Unknown'))
    for d in [train_df, test_df, df]:
        d[col + '_le'] = d[col].fillna('Unknown').apply(
            lambda x: le.transform([x])[0] if x in le.classes_ else -1
        )
    le_dict[col] = le

FEATURE_COLS = [
    'units_claimed',          
    'log_amount',            
    'amount_vs_dir_med',      
    'amount_vs_group_median', 
    'merit_volume',          
    'is_large_farm',     
    'direction_priority',    

    'risk_score_rule',        
    'amount_anomaly_flag',    
    'amount_vs_norm_median',  
    'group_z_amount',         
    'mortality_norm_pct',    
    'pasture_load_log',       
    'mortality_risk_log',     

    'direction_enc_s',       
    'normative_enc_s',        
    'direction_cnt',      
    'direction_le',       

    'eligibility_score_raw',  
    'eligible_flag',         

    'normative',

    'policy_impact_score',
    'budget_efficiency_score',
    'normative_efficiency',
]

X_train = train_df[FEATURE_COLS].fillna(0)
y_train = train_df['target']
X_test  = test_df[FEATURE_COLS].fillna(0)
y_test  = test_df['target']

print(f"  Признаков: {len(FEATURE_COLS)}")
print("  ✓ day_of_year / month / early_submission убраны из ядра")
print("  ✓ region_approval_rate заменён сглаженным (Bayesian k=50)")

print("\n▶ ШАГ 5: Обучение LightGBM")

import lightgbm as lgb
from sklearn.metrics import roc_auc_score

model = lgb.LGBMClassifier(
    n_estimators      = 600,
    learning_rate     = 0.015,
    max_depth         = 4,
    num_leaves        = 15,
    min_child_samples = 60,
    subsample         = 0.65,
    colsample_bytree  = 0.65,
    reg_alpha         = 0.5,
    reg_lambda        = 2.0,
    random_state      = 42,
    verbose           = -1,
    class_weight      = 'balanced',
)

model.fit(
    X_train, y_train,
    eval_set=[(X_test, y_test)],
    callbacks=[lgb.early_stopping(60, verbose=False), lgb.log_evaluation(period=-1)]
)

train_proba = model.predict_proba(X_train)[:, 1]
test_proba  = model.predict_proba(X_test)[:,  1]
train_auc   = roc_auc_score(y_train, train_proba)
test_auc    = roc_auc_score(y_test,  test_proba)
gap         = train_auc - test_auc

print(f"\n  ROC-AUC train: {train_auc:.4f}")
print(f"  ROC-AUC test:  {test_auc:.4f}  ← честная метрика")
print(f"  Разрыв:        {gap:.4f}  {'⚠️  переобучение' if gap > 0.05 else '✓ норма'}")

print("\n▶ ШАГ 6: SHAP — важность признаков")

import shap

explainer   = shap.TreeExplainer(model)
shap_vals   = explainer.shap_values(X_train)
shap_matrix = shap_vals[1] if isinstance(shap_vals, list) else shap_vals

feature_importance = pd.DataFrame({
    'feature':    FEATURE_COLS,
    'importance': np.abs(shap_matrix).mean(axis=0)
}).sort_values('importance', ascending=False)

print("\n  Топ-12 признаков:")
max_imp = feature_importance['importance'].max()
for _, row in feature_importance.head(12).iterrows():
    bar = "█" * int(row['importance'] * 28 / max(max_imp, 1e-9))
    print(f"  {row['feature']:32s} {bar} {row['importance']:.4f}")

regional = ['region_enc_s', 'region_le', 'region_cnt']
temporal = ['quarter', 'week_of_year']
top5     = feature_importance.head(5)['feature'].tolist()
if any(t in top5 for t in temporal):
    print("\n  ⚠️  Внимание: временны́е признаки попали в топ-5! (не должны быть в FEATURE_COLS)")
if any(r in top5 for r in regional):
    print("\n  ⚠️  Внимание: региональные прокси попали в топ-5! (не должны быть в FEATURE_COLS)")

print("\n▶ ШАГ 7: Финальный Score (0.55 ML + 0.25 Eligibility + 0.12 Risk + 0.08 BudgetEff)")

X_all = df[FEATURE_COLS].fillna(0)
df['ml_score'] = model.predict_proba(X_all)[:, 1]

df['eligibility_component'] = df['eligibility_score_raw']         
df['risk_component']        = (1 - df['risk_score_rule']).clip(0,1)
df['merit_component']       = df['ml_score']                       

df['budget_eff_component'] = df['budget_efficiency_score'].clip(0, 1)
df['final_score'] = (
    0.55 * df['merit_component'] +
    0.25 * df['eligibility_component'] +
    0.12 * df['risk_component'] +
    0.08 * df['budget_eff_component']
).clip(0, 1)

df.loc[df['eligible_flag'] == 0, 'final_score'] *= 0.35

df.loc[df['manual_review_flag'] == 1, 'final_score'] *= 0.75

df['final_score_pct']         = (df['final_score']           * 100).round(1)
df['ml_score_pct']            = (df['ml_score']              * 100).round(1)

df['eligibility_component_pct']   = (df['eligibility_component']   * 25).round(1)
df['merit_component_pct']         = (df['merit_component']         * 55).round(1)
df['risk_component_pct']          = (df['risk_component']          * 12).round(1)
df['budget_eff_component_pct']    = (df['budget_eff_component']    *  8).round(1)

print("\n▶ ШАГ 8: Рекомендации + Reason Codes")

def get_recommendation(row):
    if row['eligible_flag'] == 0:
        return 'REJECT'
    if row['npa_conflict_count'] >= 2 or row['risk_bucket'] == 'HIGH':
        return 'MANUAL_REVIEW'
    if row['manual_review_flag'] == 1:
        return 'MANUAL_REVIEW'
    return 'APPROVE'

df['recommendation'] = df.apply(get_recommendation, axis=1)

def build_reason_codes(row):
    reasons = []

    if row.get('direction_priority', 0) >= 3:
        reasons.append('HIGH_PRIORITY_DIRECTION')
    if row.get('budget_efficiency_score', 0) >= 0.80:
        reasons.append('HIGH_BUDGET_EFFICIENCY')
    if row.get('min_normative_ok', 0) == 1 and row.get('amount_vs_normative_ok', 0) == 1:
        reasons.append('NORMATIVE_MATCH')

    if row.get('amount_anomaly_flag', 0) == 1:
        reasons.append('AMOUNT_OUTLIER')
    if row.get('pasture_overload_flag', 0) == 1:
        reasons.append('PASTURE_OVERLOAD')         
    if row.get('mortality_high_flag', 0) == 1:
        reasons.append('MORTALITY_ABOVE_NPA')      
    if row.get('missing_regulatory_fields_flag', 0) == 1:
        reasons.append('REGULATORY_GAP')
    if row.get('manual_review_flag', 0) == 1:
        reasons.append('MANUAL_REVIEW_REQUIRED')

    return '; '.join(reasons) if reasons else 'STANDARD_CASE'

df['reason_codes'] = df.apply(build_reason_codes, axis=1)

rec_dist = df['recommendation'].value_counts()
print("\n  Распределение рекомендаций:")
for rec, cnt in rec_dist.items():
    print(f"    {rec:20s}: {cnt:,}")

print("\n▶ ШАГ 9: Shortlist")

active_df = df[df['is_active_application'] == 1].copy()


SHORTLIST_COLS = [
    'id', 'date', 'region', 'direction', 'district',
    'status', 'normative', 'amount', 'units_claimed',
    'eligible_flag', 'risk_score_rule', 'manual_review_flag',
    'eligibility_component_pct', 'merit_component_pct', 'risk_component_pct',
    'amount_vs_dir_med', 'amount_anomaly_flag', 'is_large_farm', 'direction_priority',
    'ml_score_pct', 'final_score_pct', 'budget_eff_component_pct',
    'recommendation', 'reason_codes',
    'legal_compliance_score', 'missing_regulatory_fields_flag',
    'npa_conflict_count', 'risk_bucket',
    'policy_impact_score', 'budget_efficiency_score', 'normative_efficiency',
    'pasture_overload_flag', 'mortality_high_flag',
    'min_normative_ok', 'min_units_ok',
]

print(f"\n  ── Диагностика пула для shortlist ──")
print(f"  Всего строк df:              {len(df):,}")
print(f"  APPROVE:                     {(df['recommendation']=='APPROVE').sum():,}")
print(f"  MANUAL_REVIEW:               {(df['recommendation']=='MANUAL_REVIEW').sum():,}")
print(f"  REJECT:                      {(df['recommendation']=='REJECT').sum():,}")
print(f"  npa_conflict_count >= 1:     {(df['npa_conflict_count'] >= 1).sum():,}")
print(f"  risk_score_rule > 0:         {(df['risk_score_rule'] > 0).sum():,}")
print(f"  manual_review_flag == 1:     {(df['manual_review_flag'] == 1).sum():,}")
print(f"  amount_anomaly_flag == 1:    {(df['amount_anomaly_flag'] == 1).sum():,}")
print(f"  pasture_overload_flag == 1:  {(df['pasture_overload_flag'] == 1).sum():,}")
print(f"  mortality_high_flag == 1:    {(df['mortality_high_flag'] == 1).sum():,}")

slot_approve = df[
    (df['recommendation'] == 'APPROVE') &
    (df['npa_conflict_count'] == 0) &
    (df['risk_score_rule'] == 0.0)
].sort_values(['final_score_pct', 'date', 'region'], ascending=[False, True, True]).head(12).copy()
slot_approve['shortlist_group'] = 'CLEAN_APPROVE'

all_used = set(slot_approve['id'].tolist())


slot_review = df[
    (df['recommendation'] == 'MANUAL_REVIEW') &
    (~df['id'].isin(all_used))
].sort_values(['ml_score_pct', 'date', 'region'], ascending=[False, True, True]).head(6).copy()


if len(slot_review) < 6:
    need = 6 - len(slot_review)
    extra = df[
        (df['manual_review_flag'] == 1) &
        (~df['id'].isin(all_used | set(slot_review['id'].tolist())))
    ].sort_values('ml_score_pct', ascending=False).head(need).copy()
    slot_review = pd.concat([slot_review, extra])


if len(slot_review) < 6:
    need = 6 - len(slot_review)
    extra = df[
        (df['npa_conflict_count'] >= 1) &
        (~df['id'].isin(all_used | set(slot_review['id'].tolist())))
    ].sort_values('ml_score_pct', ascending=False).head(need).copy()
    slot_review = pd.concat([slot_review, extra])


if len(slot_review) < 6:
    need = 6 - len(slot_review)
    extra = df[
        (df['risk_score_rule'] > 0) &
        (~df['id'].isin(all_used | set(slot_review['id'].tolist())))
    ].sort_values('risk_score_rule', ascending=False).head(need).copy()
    slot_review = pd.concat([slot_review, extra])


if len(slot_review) < 6:
    need = 6 - len(slot_review)
    extra = df[
        (
            (df['amount_anomaly_flag'] == 1) |
            (df['pasture_overload_flag'] == 1) |
            (df['mortality_high_flag'] == 1) |
            (df['missing_regulatory_fields_flag'] == 1)
        ) &
        (~df['id'].isin(all_used | set(slot_review['id'].tolist())))
    ].sort_values('ml_score_pct', ascending=False).head(need).copy()
    slot_review = pd.concat([slot_review, extra])

slot_review['shortlist_group'] = 'HIGH_MERIT_RISK'
all_used |= set(slot_review['id'].tolist())


slot_risk = df[
    (df['npa_conflict_count'] >= 1) &
    (~df['id'].isin(all_used))
].sort_values('npa_conflict_count', ascending=False).head(2).copy()

if len(slot_risk) < 2:
    need = 2 - len(slot_risk)
    extra = df[
        (df['risk_score_rule'] > 0) &
        (~df['id'].isin(all_used | set(slot_risk['id'].tolist())))
    ].sort_values('risk_score_rule', ascending=False).head(need).copy()
    slot_risk = pd.concat([slot_risk, extra])

slot_risk['shortlist_group'] = 'NPA_CONFLICT'

shortlist_combined = pd.concat(
    [slot_approve, slot_review, slot_risk], ignore_index=True
).drop_duplicates(subset='id')
shortlist_combined['rank'] = np.arange(1, len(shortlist_combined) + 1)

print(f"\n  ── Квоты shortlist ──")
print(f"  CLEAN_APPROVE   (risk=0):     {len(slot_approve)}")
print(f"  HIGH_MERIT_RISK (риск):       {len(slot_review)}")
print(f"  NPA_CONFLICT    (конфликты):  {len(slot_risk)}")


shortlist = shortlist_combined.head(20)[SHORTLIST_COLS + ['shortlist_group', 'rank']].reset_index(drop=True)
shortlist.index += 1

print(f"\n  Shortlist (топ-20):")
print(shortlist[[
    'rank', 'region', 'direction', 'units_claimed',
    'risk_score_rule', 'npa_conflict_count', 'risk_bucket',
    'final_score_pct', 'recommendation', 'shortlist_group', 'reason_codes'
]].to_string())
print(f"\n  Регионов:    {shortlist['region'].nunique()}")
print(f"  Направлений: {shortlist['direction'].nunique()}")

conflicts_in_list = shortlist[shortlist["npa_conflict_count"] >= 1]
print(f"\n  ⚡ Контрпримеры (NPA-конфликт): {len(conflicts_in_list)}")
if len(conflicts_in_list) > 0:
    print(conflicts_in_list[[
        "region","direction","npa_conflict_count","risk_bucket",
        "ml_score_pct","final_score_pct","shortlist_group","reason_codes"
    ]].to_string())

print("\n▶ ШАГ 10: Сохранение")

df.to_csv("scored_applications.csv", index=False, encoding='utf-8-sig')
print("  ✓ scored_applications.csv")

shortlist.to_csv("shortlist_top20.csv", index=True, encoding='utf-8-sig')
print("  ✓ shortlist_top20.csv")

feature_importance.to_csv("feature_importance.csv", index=False, encoding='utf-8-sig')
print("  ✓ feature_importance.csv")

approved       = int((active_df['recommendation'] == 'APPROVE').sum())
manual_review  = int((active_df['recommendation'] == 'MANUAL_REVIEW').sum())
rejected       = int((active_df['recommendation'] == 'REJECT').sum())

data_json = {
    "meta": {
        "total_applications":        int(len(df)),
        "active_applications":       int(len(active_df)),
        "eligible_applications":     int(df['eligible_flag'].sum()),
        "approved_candidates":       approved,
        "manual_review_candidates":  manual_review,
        "rejected_candidates":       rejected,
        "roc_auc_test":              round(float(test_auc),  4),
        "roc_auc_train":             round(float(train_auc), 4),
        "train_test_gap":            round(float(gap), 4),
        "shortlist_regions":         int(shortlist['region'].nunique()),
        "shortlist_directions":      int(shortlist['direction'].nunique()),
        "regions_count":             int(df['region'].nunique()),
        "directions_count":          int(df['direction'].nunique()),
        "generated_at":              date.today().isoformat(),
        "model_version":             "final",
        "score_formula":             "0.55*Merit_ML + 0.25*Eligibility + 0.12*(1-Risk_NPA) + 0.08*BudgetEfficiency",
        "npa_sources": [
            "Приказ №18404 — direction-specific eligibility rules (нормативы, поголовье)",
            "Приказ №12488 — нормы естественной убыли → mortality_deviation_ratio",
            "Приказ №11064 — пастбищная нагрузка → pasture_overload_ratio"
        ]
    },
    "shortlist": [
        {
            "rank":               int(i + 1),
            "id":                 str(row['id']),
            "date":               str(row['date'].date() if hasattr(row['date'], 'date') else row['date']),
            "region":             str(row['region'])    if pd.notna(row['region'])    else "—",
            "direction":          str(row['direction']) if pd.notna(row['direction']) else "—",
            "district":           str(row['district'])  if pd.notna(row['district'])  else "—",
            "status":             str(row['status'])    if pd.notna(row['status'])    else "—",
            "normative":          float(row['normative'])     if pd.notna(row['normative'])     else 0,
            "amount":             float(row['amount'])        if pd.notna(row['amount'])        else 0,
            "units_claimed":      float(row['units_claimed']) if pd.notna(row['units_claimed']) else 0,
            "eligible_flag":      int(row['eligible_flag']),
            "risk_score_rule":    round(float(row['risk_score_rule']), 3),
            "manual_review_flag": int(row['manual_review_flag']),
            "eligibility_component_pct": float(row['eligibility_component_pct']),  
            "budget_eff_component_pct":   float(row.get('budget_eff_component_pct', 0)), 
            "merit_component_pct":       float(row['merit_component_pct']),       
            "risk_component_pct":        float(row['risk_component_pct']),          
            "amount_vs_dir_med":  round(float(row['amount_vs_dir_med']), 3),
            "amount_anomaly_flag": int(row['amount_anomaly_flag']),
            "is_large_farm":      int(row['is_large_farm']),
            "direction_priority": int(row['direction_priority']),
            "ml_score_pct":       float(row['ml_score_pct']),
            "final_score_pct":    float(row['final_score_pct']),
            "recommendation":     str(row['recommendation']),
            "reason_codes":       str(row['reason_codes']),
            "legal_compliance_score":        round(float(row.get('legal_compliance_score', 0)), 3),
            "missing_regulatory_fields_flag": int(row.get('missing_regulatory_fields_flag', 0)),
            "npa_conflict_count":            int(row.get('npa_conflict_count', 0)),
            "risk_bucket":                   str(row.get('risk_bucket', 'LOW')),
            "policy_impact_score":           round(float(row.get('policy_impact_score', 0)), 3),
            "budget_efficiency_score":       round(float(row.get('budget_efficiency_score', 0)), 3),
            "normative_efficiency":          round(float(row.get('normative_efficiency', 0)), 3),
            "pasture_overload_flag":         int(row.get('pasture_overload_flag', 0)),
            "mortality_high_flag":           int(row.get('mortality_high_flag', 0)),
            "min_normative_ok":              int(row.get('min_normative_ok', 1)),
            "min_units_ok":                  int(row.get('min_units_ok', 1)),
            "shortlist_group":               str(row.get('shortlist_group', 'CLEAN_APPROVE')),
        }
        for i, (_, row) in enumerate(shortlist.iterrows())
    ],
    "feature_importance": [
        {
            "feature":    row['feature'],
            "importance": round(float(row['importance']), 4),
            "category": (
                "merit"       if row['feature'] in [
                    'units_claimed','log_amount','amount_vs_dir_med',
                    'merit_volume','is_large_farm','direction_priority',
                    'amount_vs_group_median'] else
                "risk_npa"    if row['feature'] in [
                    'risk_score_rule','amount_anomaly_flag',
                    'mortality_norm_pct','pasture_load_log',
                    'mortality_risk_log','group_z_amount',
                    'amount_vs_norm_median'] else
                "eligibility" if row['feature'] in [
                    'eligibility_score_raw','eligible_flag'] else
                "temporal"    if row['feature'] in ['quarter','week_of_year'] else
                "categorical"
            )
        }
        for _, row in feature_importance.iterrows()
    ],
    "directions": [
        {
            "direction": row['direction'],
            "count":     int(row['count']),
            "pct":       round(float(row['count']) / len(df) * 100, 1)
        }
        for _, row in df.groupby('direction').size()
                         .reset_index(name='count')
                         .sort_values('count', ascending=False).iterrows()
    ],
    "recommendation_distribution": {
        str(k): int(v) for k, v in df['recommendation'].value_counts().items()
    }
}

with open("data.json", "w", encoding="utf-8") as jf:
    json.dump(data_json, jf, ensure_ascii=False, indent=2, allow_nan=False)
print("  ✓ data.json")

print(f"""
╔══════════════════════════════════════════════════════════════════╗
║  ИТОГИ — ФИНАЛЬНАЯ ВЕРСИЯ                                       ║
╠══════════════════════════════════════════════════════════════════╣
║  Архитектура:   Eligibility → Risk → Merit ML (3 уровня)       ║
║  Score formula: 0.55×ML + 0.25×Elig + 0.12×Risk + 0.08×BudEff║
║  ROC-AUC train: {train_auc:.4f}                                        ║
║  ROC-AUC test:  {test_auc:.4f}  ← честная метрика                  ║
║  Разрыв:        {gap:.4f}                                        ║
╠══════════════════════════════════════════════════════════════════╣
║  Всего заявок:       {len(df):,}                                ║
║  Eligible:           {df['eligible_flag'].sum():,}                              ║
║  APPROVE:            {approved:,}                                          ║
║  MANUAL_REVIEW:      {manual_review:,}                                         ║
║  REJECT:             {rejected:,}                                           ║
╠══════════════════════════════════════════════════════════════════╣
║  НПА в пайплайне:                                               ║
║    ✓ №18404 — eligibility filter                                ║
║    ✓ №12488 — нормы падежа → risk                               ║
║    ✓ №11064 — пастбищная нагрузка → risk                        ║
╠══════════════════════════════════════════════════════════════════╣
║  Выходные файлы:                                                ║
║    scored_applications.csv                                ║
║    shortlist_top20.csv                                    ║
║    feature_importance.csv                                 ║
║    data.json  ← подключи к dashboard.html                 ║
╚══════════════════════════════════════════════════════════════════╝
""")