// rules.js — 168 treatment rules (7 volumes × 6 chlorine × 4 pH bands)
// getRecommendation(poolVolume, chlorine, ph) → {cl_granule_gr, cl_tablete, ph_kg, antialgic_l} | null

'use strict';

const GRAMS_PER_TABLET = 250;

// 168 rules: [vol_min, vol_max, cl_min, cl_max, ph_min, ph_max, cl_gr, cl_tab, ph_kg, anti_l]
// Volume bands  (m³): 30-40, 41-60, 61-80, 81-100, 101-130, 131-180, 181-200
// Chlorine bands     : 0-0.29, 0.3-0.99, 1.0-1.49, 1.5-1.99, 2.0-2.99, 3.0+
// pH bands           : 7.0-7.29, 7.3-7.59, 7.6-7.99, 8.0-8.5
// Where original had ranges (e.g. "600-700 gr"), midpoint is used.

const TREATMENT_RULES = [
  // ═══════════════════════════════════════════════
  // VOLUM: 30-40 m³
  // ═══════════════════════════════════════════════
  // Cl 0-0.29
  {pool_vol_min:30,pool_vol_max:40, cl_min:0,   cl_max:0.29, ph_min:7.0, ph_max:7.29, rec_cl_gr:650,  rec_cl_tab:3, rec_ph_kg:0,   rec_anti:0.5},
  {pool_vol_min:30,pool_vol_max:40, cl_min:0,   cl_max:0.29, ph_min:7.3, ph_max:7.59, rec_cl_gr:650,  rec_cl_tab:3, rec_ph_kg:0.5, rec_anti:0.5},
  {pool_vol_min:30,pool_vol_max:40, cl_min:0,   cl_max:0.29, ph_min:7.6, ph_max:7.99, rec_cl_gr:650,  rec_cl_tab:3, rec_ph_kg:1.0, rec_anti:0.5},
  {pool_vol_min:30,pool_vol_max:40, cl_min:0,   cl_max:0.29, ph_min:8.0, ph_max:8.5,  rec_cl_gr:650,  rec_cl_tab:3, rec_ph_kg:1.5, rec_anti:0.5},
  // Cl 0.3-0.99
  {pool_vol_min:30,pool_vol_max:40, cl_min:0.3, cl_max:0.99, ph_min:7.0, ph_max:7.29, rec_cl_gr:400,  rec_cl_tab:2, rec_ph_kg:0,   rec_anti:0.5},
  {pool_vol_min:30,pool_vol_max:40, cl_min:0.3, cl_max:0.99, ph_min:7.3, ph_max:7.59, rec_cl_gr:400,  rec_cl_tab:2, rec_ph_kg:0.4, rec_anti:0.5},
  {pool_vol_min:30,pool_vol_max:40, cl_min:0.3, cl_max:0.99, ph_min:7.6, ph_max:7.99, rec_cl_gr:400,  rec_cl_tab:2, rec_ph_kg:0.8, rec_anti:0.5},
  {pool_vol_min:30,pool_vol_max:40, cl_min:0.3, cl_max:0.99, ph_min:8.0, ph_max:8.5,  rec_cl_gr:400,  rec_cl_tab:2, rec_ph_kg:1.2, rec_anti:0.5},
  // Cl 1.0-1.49
  {pool_vol_min:30,pool_vol_max:40, cl_min:1.0, cl_max:1.49, ph_min:7.0, ph_max:7.29, rec_cl_gr:200,  rec_cl_tab:1, rec_ph_kg:0,   rec_anti:0.5},
  {pool_vol_min:30,pool_vol_max:40, cl_min:1.0, cl_max:1.49, ph_min:7.3, ph_max:7.59, rec_cl_gr:200,  rec_cl_tab:1, rec_ph_kg:0.3, rec_anti:0.5},
  {pool_vol_min:30,pool_vol_max:40, cl_min:1.0, cl_max:1.49, ph_min:7.6, ph_max:7.99, rec_cl_gr:200,  rec_cl_tab:1, rec_ph_kg:0.6, rec_anti:0.5},
  {pool_vol_min:30,pool_vol_max:40, cl_min:1.0, cl_max:1.49, ph_min:8.0, ph_max:8.5,  rec_cl_gr:200,  rec_cl_tab:1, rec_ph_kg:1.0, rec_anti:0.5},
  // Cl 1.5-1.99
  {pool_vol_min:30,pool_vol_max:40, cl_min:1.5, cl_max:1.99, ph_min:7.0, ph_max:7.29, rec_cl_gr:0,    rec_cl_tab:0, rec_ph_kg:0,   rec_anti:0.5},
  {pool_vol_min:30,pool_vol_max:40, cl_min:1.5, cl_max:1.99, ph_min:7.3, ph_max:7.59, rec_cl_gr:0,    rec_cl_tab:0, rec_ph_kg:0.3, rec_anti:0.5},
  {pool_vol_min:30,pool_vol_max:40, cl_min:1.5, cl_max:1.99, ph_min:7.6, ph_max:7.99, rec_cl_gr:0,    rec_cl_tab:0, rec_ph_kg:0.5, rec_anti:0.5},
  {pool_vol_min:30,pool_vol_max:40, cl_min:1.5, cl_max:1.99, ph_min:8.0, ph_max:8.5,  rec_cl_gr:0,    rec_cl_tab:0, rec_ph_kg:0.8, rec_anti:0.5},
  // Cl 2.0-2.99
  {pool_vol_min:30,pool_vol_max:40, cl_min:2.0, cl_max:2.99, ph_min:7.0, ph_max:7.29, rec_cl_gr:0,    rec_cl_tab:0, rec_ph_kg:0,   rec_anti:0},
  {pool_vol_min:30,pool_vol_max:40, cl_min:2.0, cl_max:2.99, ph_min:7.3, ph_max:7.59, rec_cl_gr:0,    rec_cl_tab:0, rec_ph_kg:0.2, rec_anti:0},
  {pool_vol_min:30,pool_vol_max:40, cl_min:2.0, cl_max:2.99, ph_min:7.6, ph_max:7.99, rec_cl_gr:0,    rec_cl_tab:0, rec_ph_kg:0.4, rec_anti:0},
  {pool_vol_min:30,pool_vol_max:40, cl_min:2.0, cl_max:2.99, ph_min:8.0, ph_max:8.5,  rec_cl_gr:0,    rec_cl_tab:0, rec_ph_kg:0.7, rec_anti:0},
  // Cl 3.0+
  {pool_vol_min:30,pool_vol_max:40, cl_min:3.0, cl_max:99,   ph_min:7.0, ph_max:7.29, rec_cl_gr:0,    rec_cl_tab:0, rec_ph_kg:0,   rec_anti:0},
  {pool_vol_min:30,pool_vol_max:40, cl_min:3.0, cl_max:99,   ph_min:7.3, ph_max:7.59, rec_cl_gr:0,    rec_cl_tab:0, rec_ph_kg:0,   rec_anti:0},
  {pool_vol_min:30,pool_vol_max:40, cl_min:3.0, cl_max:99,   ph_min:7.6, ph_max:7.99, rec_cl_gr:0,    rec_cl_tab:0, rec_ph_kg:0.3, rec_anti:0},
  {pool_vol_min:30,pool_vol_max:40, cl_min:3.0, cl_max:99,   ph_min:8.0, ph_max:8.5,  rec_cl_gr:0,    rec_cl_tab:0, rec_ph_kg:0.5, rec_anti:0},

  // ═══════════════════════════════════════════════
  // VOLUM: 41-60 m³
  // ═══════════════════════════════════════════════
  {pool_vol_min:41,pool_vol_max:60, cl_min:0,   cl_max:0.29, ph_min:7.0, ph_max:7.29, rec_cl_gr:900,  rec_cl_tab:4, rec_ph_kg:0,   rec_anti:0.75},
  {pool_vol_min:41,pool_vol_max:60, cl_min:0,   cl_max:0.29, ph_min:7.3, ph_max:7.59, rec_cl_gr:900,  rec_cl_tab:4, rec_ph_kg:0.75,rec_anti:0.75},
  {pool_vol_min:41,pool_vol_max:60, cl_min:0,   cl_max:0.29, ph_min:7.6, ph_max:7.99, rec_cl_gr:900,  rec_cl_tab:4, rec_ph_kg:1.5, rec_anti:0.75},
  {pool_vol_min:41,pool_vol_max:60, cl_min:0,   cl_max:0.29, ph_min:8.0, ph_max:8.5,  rec_cl_gr:900,  rec_cl_tab:4, rec_ph_kg:2.0, rec_anti:0.75},
  {pool_vol_min:41,pool_vol_max:60, cl_min:0.3, cl_max:0.99, ph_min:7.0, ph_max:7.29, rec_cl_gr:600,  rec_cl_tab:3, rec_ph_kg:0,   rec_anti:0.75},
  {pool_vol_min:41,pool_vol_max:60, cl_min:0.3, cl_max:0.99, ph_min:7.3, ph_max:7.59, rec_cl_gr:600,  rec_cl_tab:3, rec_ph_kg:0.6, rec_anti:0.75},
  {pool_vol_min:41,pool_vol_max:60, cl_min:0.3, cl_max:0.99, ph_min:7.6, ph_max:7.99, rec_cl_gr:600,  rec_cl_tab:3, rec_ph_kg:1.2, rec_anti:0.75},
  {pool_vol_min:41,pool_vol_max:60, cl_min:0.3, cl_max:0.99, ph_min:8.0, ph_max:8.5,  rec_cl_gr:600,  rec_cl_tab:3, rec_ph_kg:1.8, rec_anti:0.75},
  {pool_vol_min:41,pool_vol_max:60, cl_min:1.0, cl_max:1.49, ph_min:7.0, ph_max:7.29, rec_cl_gr:300,  rec_cl_tab:1, rec_ph_kg:0,   rec_anti:0.75},
  {pool_vol_min:41,pool_vol_max:60, cl_min:1.0, cl_max:1.49, ph_min:7.3, ph_max:7.59, rec_cl_gr:300,  rec_cl_tab:1, rec_ph_kg:0.4, rec_anti:0.75},
  {pool_vol_min:41,pool_vol_max:60, cl_min:1.0, cl_max:1.49, ph_min:7.6, ph_max:7.99, rec_cl_gr:300,  rec_cl_tab:1, rec_ph_kg:0.9, rec_anti:0.75},
  {pool_vol_min:41,pool_vol_max:60, cl_min:1.0, cl_max:1.49, ph_min:8.0, ph_max:8.5,  rec_cl_gr:300,  rec_cl_tab:1, rec_ph_kg:1.4, rec_anti:0.75},
  {pool_vol_min:41,pool_vol_max:60, cl_min:1.5, cl_max:1.99, ph_min:7.0, ph_max:7.29, rec_cl_gr:0,    rec_cl_tab:0, rec_ph_kg:0,   rec_anti:0.5},
  {pool_vol_min:41,pool_vol_max:60, cl_min:1.5, cl_max:1.99, ph_min:7.3, ph_max:7.59, rec_cl_gr:0,    rec_cl_tab:0, rec_ph_kg:0.4, rec_anti:0.5},
  {pool_vol_min:41,pool_vol_max:60, cl_min:1.5, cl_max:1.99, ph_min:7.6, ph_max:7.99, rec_cl_gr:0,    rec_cl_tab:0, rec_ph_kg:0.7, rec_anti:0.5},
  {pool_vol_min:41,pool_vol_max:60, cl_min:1.5, cl_max:1.99, ph_min:8.0, ph_max:8.5,  rec_cl_gr:0,    rec_cl_tab:0, rec_ph_kg:1.1, rec_anti:0.5},
  {pool_vol_min:41,pool_vol_max:60, cl_min:2.0, cl_max:2.99, ph_min:7.0, ph_max:7.29, rec_cl_gr:0,    rec_cl_tab:0, rec_ph_kg:0,   rec_anti:0},
  {pool_vol_min:41,pool_vol_max:60, cl_min:2.0, cl_max:2.99, ph_min:7.3, ph_max:7.59, rec_cl_gr:0,    rec_cl_tab:0, rec_ph_kg:0.3, rec_anti:0},
  {pool_vol_min:41,pool_vol_max:60, cl_min:2.0, cl_max:2.99, ph_min:7.6, ph_max:7.99, rec_cl_gr:0,    rec_cl_tab:0, rec_ph_kg:0.6, rec_anti:0},
  {pool_vol_min:41,pool_vol_max:60, cl_min:2.0, cl_max:2.99, ph_min:8.0, ph_max:8.5,  rec_cl_gr:0,    rec_cl_tab:0, rec_ph_kg:1.0, rec_anti:0},
  {pool_vol_min:41,pool_vol_max:60, cl_min:3.0, cl_max:99,   ph_min:7.0, ph_max:7.29, rec_cl_gr:0,    rec_cl_tab:0, rec_ph_kg:0,   rec_anti:0},
  {pool_vol_min:41,pool_vol_max:60, cl_min:3.0, cl_max:99,   ph_min:7.3, ph_max:7.59, rec_cl_gr:0,    rec_cl_tab:0, rec_ph_kg:0,   rec_anti:0},
  {pool_vol_min:41,pool_vol_max:60, cl_min:3.0, cl_max:99,   ph_min:7.6, ph_max:7.99, rec_cl_gr:0,    rec_cl_tab:0, rec_ph_kg:0.4, rec_anti:0},
  {pool_vol_min:41,pool_vol_max:60, cl_min:3.0, cl_max:99,   ph_min:8.0, ph_max:8.5,  rec_cl_gr:0,    rec_cl_tab:0, rec_ph_kg:0.7, rec_anti:0},

  // ═══════════════════════════════════════════════
  // VOLUM: 61-80 m³
  // ═══════════════════════════════════════════════
  {pool_vol_min:61,pool_vol_max:80, cl_min:0,   cl_max:0.29, ph_min:7.0, ph_max:7.29, rec_cl_gr:1200, rec_cl_tab:5, rec_ph_kg:0,   rec_anti:1.0},
  {pool_vol_min:61,pool_vol_max:80, cl_min:0,   cl_max:0.29, ph_min:7.3, ph_max:7.59, rec_cl_gr:1200, rec_cl_tab:5, rec_ph_kg:1.0, rec_anti:1.0},
  {pool_vol_min:61,pool_vol_max:80, cl_min:0,   cl_max:0.29, ph_min:7.6, ph_max:7.99, rec_cl_gr:1200, rec_cl_tab:5, rec_ph_kg:2.0, rec_anti:1.0},
  {pool_vol_min:61,pool_vol_max:80, cl_min:0,   cl_max:0.29, ph_min:8.0, ph_max:8.5,  rec_cl_gr:1200, rec_cl_tab:5, rec_ph_kg:2.5, rec_anti:1.0},
  {pool_vol_min:61,pool_vol_max:80, cl_min:0.3, cl_max:0.99, ph_min:7.0, ph_max:7.29, rec_cl_gr:800,  rec_cl_tab:3, rec_ph_kg:0,   rec_anti:1.0},
  {pool_vol_min:61,pool_vol_max:80, cl_min:0.3, cl_max:0.99, ph_min:7.3, ph_max:7.59, rec_cl_gr:800,  rec_cl_tab:3, rec_ph_kg:0.8, rec_anti:1.0},
  {pool_vol_min:61,pool_vol_max:80, cl_min:0.3, cl_max:0.99, ph_min:7.6, ph_max:7.99, rec_cl_gr:800,  rec_cl_tab:3, rec_ph_kg:1.6, rec_anti:1.0},
  {pool_vol_min:61,pool_vol_max:80, cl_min:0.3, cl_max:0.99, ph_min:8.0, ph_max:8.5,  rec_cl_gr:800,  rec_cl_tab:3, rec_ph_kg:2.2, rec_anti:1.0},
  {pool_vol_min:61,pool_vol_max:80, cl_min:1.0, cl_max:1.49, ph_min:7.0, ph_max:7.29, rec_cl_gr:400,  rec_cl_tab:2, rec_ph_kg:0,   rec_anti:1.0},
  {pool_vol_min:61,pool_vol_max:80, cl_min:1.0, cl_max:1.49, ph_min:7.3, ph_max:7.59, rec_cl_gr:400,  rec_cl_tab:2, rec_ph_kg:0.5, rec_anti:1.0},
  {pool_vol_min:61,pool_vol_max:80, cl_min:1.0, cl_max:1.49, ph_min:7.6, ph_max:7.99, rec_cl_gr:400,  rec_cl_tab:2, rec_ph_kg:1.2, rec_anti:1.0},
  {pool_vol_min:61,pool_vol_max:80, cl_min:1.0, cl_max:1.49, ph_min:8.0, ph_max:8.5,  rec_cl_gr:400,  rec_cl_tab:2, rec_ph_kg:1.8, rec_anti:1.0},
  {pool_vol_min:61,pool_vol_max:80, cl_min:1.5, cl_max:1.99, ph_min:7.0, ph_max:7.29, rec_cl_gr:0,    rec_cl_tab:0, rec_ph_kg:0,   rec_anti:0.5},
  {pool_vol_min:61,pool_vol_max:80, cl_min:1.5, cl_max:1.99, ph_min:7.3, ph_max:7.59, rec_cl_gr:0,    rec_cl_tab:0, rec_ph_kg:0.5, rec_anti:0.5},
  {pool_vol_min:61,pool_vol_max:80, cl_min:1.5, cl_max:1.99, ph_min:7.6, ph_max:7.99, rec_cl_gr:0,    rec_cl_tab:0, rec_ph_kg:0.9, rec_anti:0.5},
  {pool_vol_min:61,pool_vol_max:80, cl_min:1.5, cl_max:1.99, ph_min:8.0, ph_max:8.5,  rec_cl_gr:0,    rec_cl_tab:0, rec_ph_kg:1.4, rec_anti:0.5},
  {pool_vol_min:61,pool_vol_max:80, cl_min:2.0, cl_max:2.99, ph_min:7.0, ph_max:7.29, rec_cl_gr:0,    rec_cl_tab:0, rec_ph_kg:0,   rec_anti:0},
  {pool_vol_min:61,pool_vol_max:80, cl_min:2.0, cl_max:2.99, ph_min:7.3, ph_max:7.59, rec_cl_gr:0,    rec_cl_tab:0, rec_ph_kg:0.4, rec_anti:0},
  {pool_vol_min:61,pool_vol_max:80, cl_min:2.0, cl_max:2.99, ph_min:7.6, ph_max:7.99, rec_cl_gr:0,    rec_cl_tab:0, rec_ph_kg:0.8, rec_anti:0},
  {pool_vol_min:61,pool_vol_max:80, cl_min:2.0, cl_max:2.99, ph_min:8.0, ph_max:8.5,  rec_cl_gr:0,    rec_cl_tab:0, rec_ph_kg:1.2, rec_anti:0},
  {pool_vol_min:61,pool_vol_max:80, cl_min:3.0, cl_max:99,   ph_min:7.0, ph_max:7.29, rec_cl_gr:0,    rec_cl_tab:0, rec_ph_kg:0,   rec_anti:0},
  {pool_vol_min:61,pool_vol_max:80, cl_min:3.0, cl_max:99,   ph_min:7.3, ph_max:7.59, rec_cl_gr:0,    rec_cl_tab:0, rec_ph_kg:0,   rec_anti:0},
  {pool_vol_min:61,pool_vol_max:80, cl_min:3.0, cl_max:99,   ph_min:7.6, ph_max:7.99, rec_cl_gr:0,    rec_cl_tab:0, rec_ph_kg:0.5, rec_anti:0},
  {pool_vol_min:61,pool_vol_max:80, cl_min:3.0, cl_max:99,   ph_min:8.0, ph_max:8.5,  rec_cl_gr:0,    rec_cl_tab:0, rec_ph_kg:0.9, rec_anti:0},

  // ═══════════════════════════════════════════════
  // VOLUM: 81-100 m³
  // ═══════════════════════════════════════════════
  {pool_vol_min:81,pool_vol_max:100, cl_min:0,   cl_max:0.29, ph_min:7.0, ph_max:7.29, rec_cl_gr:1600, rec_cl_tab:6, rec_ph_kg:0,   rec_anti:1.5},
  {pool_vol_min:81,pool_vol_max:100, cl_min:0,   cl_max:0.29, ph_min:7.3, ph_max:7.59, rec_cl_gr:1600, rec_cl_tab:6, rec_ph_kg:1.25,rec_anti:1.5},
  {pool_vol_min:81,pool_vol_max:100, cl_min:0,   cl_max:0.29, ph_min:7.6, ph_max:7.99, rec_cl_gr:1600, rec_cl_tab:6, rec_ph_kg:2.5, rec_anti:1.5},
  {pool_vol_min:81,pool_vol_max:100, cl_min:0,   cl_max:0.29, ph_min:8.0, ph_max:8.5,  rec_cl_gr:1600, rec_cl_tab:6, rec_ph_kg:3.5, rec_anti:1.5},
  {pool_vol_min:81,pool_vol_max:100, cl_min:0.3, cl_max:0.99, ph_min:7.0, ph_max:7.29, rec_cl_gr:1000, rec_cl_tab:4, rec_ph_kg:0,   rec_anti:1.5},
  {pool_vol_min:81,pool_vol_max:100, cl_min:0.3, cl_max:0.99, ph_min:7.3, ph_max:7.59, rec_cl_gr:1000, rec_cl_tab:4, rec_ph_kg:1.0, rec_anti:1.5},
  {pool_vol_min:81,pool_vol_max:100, cl_min:0.3, cl_max:0.99, ph_min:7.6, ph_max:7.99, rec_cl_gr:1000, rec_cl_tab:4, rec_ph_kg:2.0, rec_anti:1.5},
  {pool_vol_min:81,pool_vol_max:100, cl_min:0.3, cl_max:0.99, ph_min:8.0, ph_max:8.5,  rec_cl_gr:1000, rec_cl_tab:4, rec_ph_kg:2.8, rec_anti:1.5},
  {pool_vol_min:81,pool_vol_max:100, cl_min:1.0, cl_max:1.49, ph_min:7.0, ph_max:7.29, rec_cl_gr:500,  rec_cl_tab:2, rec_ph_kg:0,   rec_anti:1.0},
  {pool_vol_min:81,pool_vol_max:100, cl_min:1.0, cl_max:1.49, ph_min:7.3, ph_max:7.59, rec_cl_gr:500,  rec_cl_tab:2, rec_ph_kg:0.7, rec_anti:1.0},
  {pool_vol_min:81,pool_vol_max:100, cl_min:1.0, cl_max:1.49, ph_min:7.6, ph_max:7.99, rec_cl_gr:500,  rec_cl_tab:2, rec_ph_kg:1.5, rec_anti:1.0},
  {pool_vol_min:81,pool_vol_max:100, cl_min:1.0, cl_max:1.49, ph_min:8.0, ph_max:8.5,  rec_cl_gr:500,  rec_cl_tab:2, rec_ph_kg:2.2, rec_anti:1.0},
  {pool_vol_min:81,pool_vol_max:100, cl_min:1.5, cl_max:1.99, ph_min:7.0, ph_max:7.29, rec_cl_gr:0,    rec_cl_tab:0, rec_ph_kg:0,   rec_anti:0.5},
  {pool_vol_min:81,pool_vol_max:100, cl_min:1.5, cl_max:1.99, ph_min:7.3, ph_max:7.59, rec_cl_gr:0,    rec_cl_tab:0, rec_ph_kg:0.6, rec_anti:0.5},
  {pool_vol_min:81,pool_vol_max:100, cl_min:1.5, cl_max:1.99, ph_min:7.6, ph_max:7.99, rec_cl_gr:0,    rec_cl_tab:0, rec_ph_kg:1.2, rec_anti:0.5},
  {pool_vol_min:81,pool_vol_max:100, cl_min:1.5, cl_max:1.99, ph_min:8.0, ph_max:8.5,  rec_cl_gr:0,    rec_cl_tab:0, rec_ph_kg:1.8, rec_anti:0.5},
  {pool_vol_min:81,pool_vol_max:100, cl_min:2.0, cl_max:2.99, ph_min:7.0, ph_max:7.29, rec_cl_gr:0,    rec_cl_tab:0, rec_ph_kg:0,   rec_anti:0},
  {pool_vol_min:81,pool_vol_max:100, cl_min:2.0, cl_max:2.99, ph_min:7.3, ph_max:7.59, rec_cl_gr:0,    rec_cl_tab:0, rec_ph_kg:0.5, rec_anti:0},
  {pool_vol_min:81,pool_vol_max:100, cl_min:2.0, cl_max:2.99, ph_min:7.6, ph_max:7.99, rec_cl_gr:0,    rec_cl_tab:0, rec_ph_kg:1.0, rec_anti:0},
  {pool_vol_min:81,pool_vol_max:100, cl_min:2.0, cl_max:2.99, ph_min:8.0, ph_max:8.5,  rec_cl_gr:0,    rec_cl_tab:0, rec_ph_kg:1.5, rec_anti:0},
  {pool_vol_min:81,pool_vol_max:100, cl_min:3.0, cl_max:99,   ph_min:7.0, ph_max:7.29, rec_cl_gr:0,    rec_cl_tab:0, rec_ph_kg:0,   rec_anti:0},
  {pool_vol_min:81,pool_vol_max:100, cl_min:3.0, cl_max:99,   ph_min:7.3, ph_max:7.59, rec_cl_gr:0,    rec_cl_tab:0, rec_ph_kg:0,   rec_anti:0},
  {pool_vol_min:81,pool_vol_max:100, cl_min:3.0, cl_max:99,   ph_min:7.6, ph_max:7.99, rec_cl_gr:0,    rec_cl_tab:0, rec_ph_kg:0.6, rec_anti:0},
  {pool_vol_min:81,pool_vol_max:100, cl_min:3.0, cl_max:99,   ph_min:8.0, ph_max:8.5,  rec_cl_gr:0,    rec_cl_tab:0, rec_ph_kg:1.1, rec_anti:0},

  // ═══════════════════════════════════════════════
  // VOLUM: 101-130 m³
  // ═══════════════════════════════════════════════
  {pool_vol_min:101,pool_vol_max:130, cl_min:0,   cl_max:0.29, ph_min:7.0, ph_max:7.29, rec_cl_gr:2000, rec_cl_tab:8, rec_ph_kg:0,   rec_anti:2.0},
  {pool_vol_min:101,pool_vol_max:130, cl_min:0,   cl_max:0.29, ph_min:7.3, ph_max:7.59, rec_cl_gr:2000, rec_cl_tab:8, rec_ph_kg:1.5, rec_anti:2.0},
  {pool_vol_min:101,pool_vol_max:130, cl_min:0,   cl_max:0.29, ph_min:7.6, ph_max:7.99, rec_cl_gr:2000, rec_cl_tab:8, rec_ph_kg:3.0, rec_anti:2.0},
  {pool_vol_min:101,pool_vol_max:130, cl_min:0,   cl_max:0.29, ph_min:8.0, ph_max:8.5,  rec_cl_gr:2000, rec_cl_tab:8, rec_ph_kg:4.5, rec_anti:2.0},
  {pool_vol_min:101,pool_vol_max:130, cl_min:0.3, cl_max:0.99, ph_min:7.0, ph_max:7.29, rec_cl_gr:1300, rec_cl_tab:5, rec_ph_kg:0,   rec_anti:2.0},
  {pool_vol_min:101,pool_vol_max:130, cl_min:0.3, cl_max:0.99, ph_min:7.3, ph_max:7.59, rec_cl_gr:1300, rec_cl_tab:5, rec_ph_kg:1.25,rec_anti:2.0},
  {pool_vol_min:101,pool_vol_max:130, cl_min:0.3, cl_max:0.99, ph_min:7.6, ph_max:7.99, rec_cl_gr:1300, rec_cl_tab:5, rec_ph_kg:2.5, rec_anti:2.0},
  {pool_vol_min:101,pool_vol_max:130, cl_min:0.3, cl_max:0.99, ph_min:8.0, ph_max:8.5,  rec_cl_gr:1300, rec_cl_tab:5, rec_ph_kg:3.5, rec_anti:2.0},
  {pool_vol_min:101,pool_vol_max:130, cl_min:1.0, cl_max:1.49, ph_min:7.0, ph_max:7.29, rec_cl_gr:650,  rec_cl_tab:3, rec_ph_kg:0,   rec_anti:1.5},
  {pool_vol_min:101,pool_vol_max:130, cl_min:1.0, cl_max:1.49, ph_min:7.3, ph_max:7.59, rec_cl_gr:650,  rec_cl_tab:3, rec_ph_kg:0.9, rec_anti:1.5},
  {pool_vol_min:101,pool_vol_max:130, cl_min:1.0, cl_max:1.49, ph_min:7.6, ph_max:7.99, rec_cl_gr:650,  rec_cl_tab:3, rec_ph_kg:1.8, rec_anti:1.5},
  {pool_vol_min:101,pool_vol_max:130, cl_min:1.0, cl_max:1.49, ph_min:8.0, ph_max:8.5,  rec_cl_gr:650,  rec_cl_tab:3, rec_ph_kg:2.8, rec_anti:1.5},
  {pool_vol_min:101,pool_vol_max:130, cl_min:1.5, cl_max:1.99, ph_min:7.0, ph_max:7.29, rec_cl_gr:0,    rec_cl_tab:0, rec_ph_kg:0,   rec_anti:1.0},
  {pool_vol_min:101,pool_vol_max:130, cl_min:1.5, cl_max:1.99, ph_min:7.3, ph_max:7.59, rec_cl_gr:0,    rec_cl_tab:0, rec_ph_kg:0.8, rec_anti:1.0},
  {pool_vol_min:101,pool_vol_max:130, cl_min:1.5, cl_max:1.99, ph_min:7.6, ph_max:7.99, rec_cl_gr:0,    rec_cl_tab:0, rec_ph_kg:1.5, rec_anti:1.0},
  {pool_vol_min:101,pool_vol_max:130, cl_min:1.5, cl_max:1.99, ph_min:8.0, ph_max:8.5,  rec_cl_gr:0,    rec_cl_tab:0, rec_ph_kg:2.2, rec_anti:1.0},
  {pool_vol_min:101,pool_vol_max:130, cl_min:2.0, cl_max:2.99, ph_min:7.0, ph_max:7.29, rec_cl_gr:0,    rec_cl_tab:0, rec_ph_kg:0,   rec_anti:0},
  {pool_vol_min:101,pool_vol_max:130, cl_min:2.0, cl_max:2.99, ph_min:7.3, ph_max:7.59, rec_cl_gr:0,    rec_cl_tab:0, rec_ph_kg:0.7, rec_anti:0},
  {pool_vol_min:101,pool_vol_max:130, cl_min:2.0, cl_max:2.99, ph_min:7.6, ph_max:7.99, rec_cl_gr:0,    rec_cl_tab:0, rec_ph_kg:1.3, rec_anti:0},
  {pool_vol_min:101,pool_vol_max:130, cl_min:2.0, cl_max:2.99, ph_min:8.0, ph_max:8.5,  rec_cl_gr:0,    rec_cl_tab:0, rec_ph_kg:1.9, rec_anti:0},
  {pool_vol_min:101,pool_vol_max:130, cl_min:3.0, cl_max:99,   ph_min:7.0, ph_max:7.29, rec_cl_gr:0,    rec_cl_tab:0, rec_ph_kg:0,   rec_anti:0},
  {pool_vol_min:101,pool_vol_max:130, cl_min:3.0, cl_max:99,   ph_min:7.3, ph_max:7.59, rec_cl_gr:0,    rec_cl_tab:0, rec_ph_kg:0,   rec_anti:0},
  {pool_vol_min:101,pool_vol_max:130, cl_min:3.0, cl_max:99,   ph_min:7.6, ph_max:7.99, rec_cl_gr:0,    rec_cl_tab:0, rec_ph_kg:0.8, rec_anti:0},
  {pool_vol_min:101,pool_vol_max:130, cl_min:3.0, cl_max:99,   ph_min:8.0, ph_max:8.5,  rec_cl_gr:0,    rec_cl_tab:0, rec_ph_kg:1.4, rec_anti:0},

  // ═══════════════════════════════════════════════
  // VOLUM: 131-180 m³
  // ═══════════════════════════════════════════════
  {pool_vol_min:131,pool_vol_max:180, cl_min:0,   cl_max:0.29, ph_min:7.0, ph_max:7.29, rec_cl_gr:2750, rec_cl_tab:11, rec_ph_kg:0,   rec_anti:2.5},
  {pool_vol_min:131,pool_vol_max:180, cl_min:0,   cl_max:0.29, ph_min:7.3, ph_max:7.59, rec_cl_gr:2750, rec_cl_tab:11, rec_ph_kg:2.0, rec_anti:2.5},
  {pool_vol_min:131,pool_vol_max:180, cl_min:0,   cl_max:0.29, ph_min:7.6, ph_max:7.99, rec_cl_gr:2750, rec_cl_tab:11, rec_ph_kg:4.0, rec_anti:2.5},
  {pool_vol_min:131,pool_vol_max:180, cl_min:0,   cl_max:0.29, ph_min:8.0, ph_max:8.5,  rec_cl_gr:2750, rec_cl_tab:11, rec_ph_kg:6.0, rec_anti:2.5},
  {pool_vol_min:131,pool_vol_max:180, cl_min:0.3, cl_max:0.99, ph_min:7.0, ph_max:7.29, rec_cl_gr:1750, rec_cl_tab:7,  rec_ph_kg:0,   rec_anti:2.5},
  {pool_vol_min:131,pool_vol_max:180, cl_min:0.3, cl_max:0.99, ph_min:7.3, ph_max:7.59, rec_cl_gr:1750, rec_cl_tab:7,  rec_ph_kg:1.75,rec_anti:2.5},
  {pool_vol_min:131,pool_vol_max:180, cl_min:0.3, cl_max:0.99, ph_min:7.6, ph_max:7.99, rec_cl_gr:1750, rec_cl_tab:7,  rec_ph_kg:3.5, rec_anti:2.5},
  {pool_vol_min:131,pool_vol_max:180, cl_min:0.3, cl_max:0.99, ph_min:8.0, ph_max:8.5,  rec_cl_gr:1750, rec_cl_tab:7,  rec_ph_kg:5.0, rec_anti:2.5},
  {pool_vol_min:131,pool_vol_max:180, cl_min:1.0, cl_max:1.49, ph_min:7.0, ph_max:7.29, rec_cl_gr:875,  rec_cl_tab:4,  rec_ph_kg:0,   rec_anti:2.0},
  {pool_vol_min:131,pool_vol_max:180, cl_min:1.0, cl_max:1.49, ph_min:7.3, ph_max:7.59, rec_cl_gr:875,  rec_cl_tab:4,  rec_ph_kg:1.25,rec_anti:2.0},
  {pool_vol_min:131,pool_vol_max:180, cl_min:1.0, cl_max:1.49, ph_min:7.6, ph_max:7.99, rec_cl_gr:875,  rec_cl_tab:4,  rec_ph_kg:2.5, rec_anti:2.0},
  {pool_vol_min:131,pool_vol_max:180, cl_min:1.0, cl_max:1.49, ph_min:8.0, ph_max:8.5,  rec_cl_gr:875,  rec_cl_tab:4,  rec_ph_kg:3.5, rec_anti:2.0},
  {pool_vol_min:131,pool_vol_max:180, cl_min:1.5, cl_max:1.99, ph_min:7.0, ph_max:7.29, rec_cl_gr:0,    rec_cl_tab:0,  rec_ph_kg:0,   rec_anti:1.5},
  {pool_vol_min:131,pool_vol_max:180, cl_min:1.5, cl_max:1.99, ph_min:7.3, ph_max:7.59, rec_cl_gr:0,    rec_cl_tab:0,  rec_ph_kg:1.0, rec_anti:1.5},
  {pool_vol_min:131,pool_vol_max:180, cl_min:1.5, cl_max:1.99, ph_min:7.6, ph_max:7.99, rec_cl_gr:0,    rec_cl_tab:0,  rec_ph_kg:2.0, rec_anti:1.5},
  {pool_vol_min:131,pool_vol_max:180, cl_min:1.5, cl_max:1.99, ph_min:8.0, ph_max:8.5,  rec_cl_gr:0,    rec_cl_tab:0,  rec_ph_kg:3.0, rec_anti:1.5},
  {pool_vol_min:131,pool_vol_max:180, cl_min:2.0, cl_max:2.99, ph_min:7.0, ph_max:7.29, rec_cl_gr:0,    rec_cl_tab:0,  rec_ph_kg:0,   rec_anti:0},
  {pool_vol_min:131,pool_vol_max:180, cl_min:2.0, cl_max:2.99, ph_min:7.3, ph_max:7.59, rec_cl_gr:0,    rec_cl_tab:0,  rec_ph_kg:0.9, rec_anti:0},
  {pool_vol_min:131,pool_vol_max:180, cl_min:2.0, cl_max:2.99, ph_min:7.6, ph_max:7.99, rec_cl_gr:0,    rec_cl_tab:0,  rec_ph_kg:1.7, rec_anti:0},
  {pool_vol_min:131,pool_vol_max:180, cl_min:2.0, cl_max:2.99, ph_min:8.0, ph_max:8.5,  rec_cl_gr:0,    rec_cl_tab:0,  rec_ph_kg:2.5, rec_anti:0},
  {pool_vol_min:131,pool_vol_max:180, cl_min:3.0, cl_max:99,   ph_min:7.0, ph_max:7.29, rec_cl_gr:0,    rec_cl_tab:0,  rec_ph_kg:0,   rec_anti:0},
  {pool_vol_min:131,pool_vol_max:180, cl_min:3.0, cl_max:99,   ph_min:7.3, ph_max:7.59, rec_cl_gr:0,    rec_cl_tab:0,  rec_ph_kg:0,   rec_anti:0},
  {pool_vol_min:131,pool_vol_max:180, cl_min:3.0, cl_max:99,   ph_min:7.6, ph_max:7.99, rec_cl_gr:0,    rec_cl_tab:0,  rec_ph_kg:1.0, rec_anti:0},
  {pool_vol_min:131,pool_vol_max:180, cl_min:3.0, cl_max:99,   ph_min:8.0, ph_max:8.5,  rec_cl_gr:0,    rec_cl_tab:0,  rec_ph_kg:1.7, rec_anti:0},

  // ═══════════════════════════════════════════════
  // VOLUM: 181-200 m³
  // ═══════════════════════════════════════════════
  {pool_vol_min:181,pool_vol_max:200, cl_min:0,   cl_max:0.29, ph_min:7.0, ph_max:7.29, rec_cl_gr:3500, rec_cl_tab:14, rec_ph_kg:0,   rec_anti:3.0},
  {pool_vol_min:181,pool_vol_max:200, cl_min:0,   cl_max:0.29, ph_min:7.3, ph_max:7.59, rec_cl_gr:3500, rec_cl_tab:14, rec_ph_kg:2.5, rec_anti:3.0},
  {pool_vol_min:181,pool_vol_max:200, cl_min:0,   cl_max:0.29, ph_min:7.6, ph_max:7.99, rec_cl_gr:3500, rec_cl_tab:14, rec_ph_kg:5.0, rec_anti:3.0},
  {pool_vol_min:181,pool_vol_max:200, cl_min:0,   cl_max:0.29, ph_min:8.0, ph_max:8.5,  rec_cl_gr:3500, rec_cl_tab:14, rec_ph_kg:7.0, rec_anti:3.0},
  {pool_vol_min:181,pool_vol_max:200, cl_min:0.3, cl_max:0.99, ph_min:7.0, ph_max:7.29, rec_cl_gr:2200, rec_cl_tab:9,  rec_ph_kg:0,   rec_anti:3.0},
  {pool_vol_min:181,pool_vol_max:200, cl_min:0.3, cl_max:0.99, ph_min:7.3, ph_max:7.59, rec_cl_gr:2200, rec_cl_tab:9,  rec_ph_kg:2.0, rec_anti:3.0},
  {pool_vol_min:181,pool_vol_max:200, cl_min:0.3, cl_max:0.99, ph_min:7.6, ph_max:7.99, rec_cl_gr:2200, rec_cl_tab:9,  rec_ph_kg:4.0, rec_anti:3.0},
  {pool_vol_min:181,pool_vol_max:200, cl_min:0.3, cl_max:0.99, ph_min:8.0, ph_max:8.5,  rec_cl_gr:2200, rec_cl_tab:9,  rec_ph_kg:6.0, rec_anti:3.0},
  {pool_vol_min:181,pool_vol_max:200, cl_min:1.0, cl_max:1.49, ph_min:7.0, ph_max:7.29, rec_cl_gr:1100, rec_cl_tab:4,  rec_ph_kg:0,   rec_anti:2.5},
  {pool_vol_min:181,pool_vol_max:200, cl_min:1.0, cl_max:1.49, ph_min:7.3, ph_max:7.59, rec_cl_gr:1100, rec_cl_tab:4,  rec_ph_kg:1.5, rec_anti:2.5},
  {pool_vol_min:181,pool_vol_max:200, cl_min:1.0, cl_max:1.49, ph_min:7.6, ph_max:7.99, rec_cl_gr:1100, rec_cl_tab:4,  rec_ph_kg:3.0, rec_anti:2.5},
  {pool_vol_min:181,pool_vol_max:200, cl_min:1.0, cl_max:1.49, ph_min:8.0, ph_max:8.5,  rec_cl_gr:1100, rec_cl_tab:4,  rec_ph_kg:4.5, rec_anti:2.5},
  {pool_vol_min:181,pool_vol_max:200, cl_min:1.5, cl_max:1.99, ph_min:7.0, ph_max:7.29, rec_cl_gr:0,    rec_cl_tab:0,  rec_ph_kg:0,   rec_anti:2.0},
  {pool_vol_min:181,pool_vol_max:200, cl_min:1.5, cl_max:1.99, ph_min:7.3, ph_max:7.59, rec_cl_gr:0,    rec_cl_tab:0,  rec_ph_kg:1.25,rec_anti:2.0},
  {pool_vol_min:181,pool_vol_max:200, cl_min:1.5, cl_max:1.99, ph_min:7.6, ph_max:7.99, rec_cl_gr:0,    rec_cl_tab:0,  rec_ph_kg:2.5, rec_anti:2.0},
  {pool_vol_min:181,pool_vol_max:200, cl_min:1.5, cl_max:1.99, ph_min:8.0, ph_max:8.5,  rec_cl_gr:0,    rec_cl_tab:0,  rec_ph_kg:3.5, rec_anti:2.0},
  {pool_vol_min:181,pool_vol_max:200, cl_min:2.0, cl_max:2.99, ph_min:7.0, ph_max:7.29, rec_cl_gr:0,    rec_cl_tab:0,  rec_ph_kg:0,   rec_anti:0},
  {pool_vol_min:181,pool_vol_max:200, cl_min:2.0, cl_max:2.99, ph_min:7.3, ph_max:7.59, rec_cl_gr:0,    rec_cl_tab:0,  rec_ph_kg:1.0, rec_anti:0},
  {pool_vol_min:181,pool_vol_max:200, cl_min:2.0, cl_max:2.99, ph_min:7.6, ph_max:7.99, rec_cl_gr:0,    rec_cl_tab:0,  rec_ph_kg:2.0, rec_anti:0},
  {pool_vol_min:181,pool_vol_max:200, cl_min:2.0, cl_max:2.99, ph_min:8.0, ph_max:8.5,  rec_cl_gr:0,    rec_cl_tab:0,  rec_ph_kg:3.0, rec_anti:0},
  {pool_vol_min:181,pool_vol_max:200, cl_min:3.0, cl_max:99,   ph_min:7.0, ph_max:7.29, rec_cl_gr:0,    rec_cl_tab:0,  rec_ph_kg:0,   rec_anti:0},
  {pool_vol_min:181,pool_vol_max:200, cl_min:3.0, cl_max:99,   ph_min:7.3, ph_max:7.59, rec_cl_gr:0,    rec_cl_tab:0,  rec_ph_kg:0,   rec_anti:0},
  {pool_vol_min:181,pool_vol_max:200, cl_min:3.0, cl_max:99,   ph_min:7.6, ph_max:7.99, rec_cl_gr:0,    rec_cl_tab:0,  rec_ph_kg:1.2, rec_anti:0},
  {pool_vol_min:181,pool_vol_max:200, cl_min:3.0, cl_max:99,   ph_min:8.0, ph_max:8.5,  rec_cl_gr:0,    rec_cl_tab:0,  rec_ph_kg:2.0, rec_anti:0},
];

/**
 * Get treatment recommendation for given pool parameters.
 * @param {number} poolVolume  - pool volume in m³
 * @param {number} chlorine    - measured chlorine (mg/L)
 * @param {number} ph          - measured pH
 * @param {Array}  [rules]     - optional rules array (defaults to TREATMENT_RULES)
 * @returns {{ cl_granule_gr: number, cl_tablete: number, ph_kg: number, antialgic_l: number } | null}
 */
function getRecommendation(poolVolume, chlorine, ph, rules) {
  const r = rules || TREATMENT_RULES;

  // Validate inputs
  if (!poolVolume || poolVolume <= 0 || chlorine == null || ph == null) return null;

  // Clamp values to rule boundaries for out-of-range inputs
  // Volume: rules cover 30-200 m³; clamp then scale proportionally
  var vol = poolVolume;
  if (vol < 30) vol = 30;
  if (vol > 200) vol = 200;

  // pH: rules cover 7.0-8.5; clamp to nearest band
  var phC = ph;
  if (phC < 7.0) phC = 7.0;
  if (phC > 8.5) phC = 8.5;

  // Chlorine: rules cover 0-99; no clamping needed (3.0+ band goes to 99)
  var clC = Math.max(0, chlorine);

  const rule = r.find(rule =>
    vol  >= rule.pool_vol_min && vol  <= rule.pool_vol_max &&
    clC  >= rule.cl_min       && clC  <= rule.cl_max &&
    phC  >= rule.ph_min       && phC  <= rule.ph_max
  );
  if (!rule) return null;

  var result = {
    cl_granule_gr: rule.rec_cl_gr,
    cl_tablete:    rule.rec_cl_tab,
    ph_kg:         rule.rec_ph_kg,
    antialgic_l:   rule.rec_anti
  };

  // Scale proportionally if volume was clamped (outside 30-200 range)
  if (poolVolume !== vol) {
    var scale = poolVolume / vol;
    result.cl_granule_gr = Math.round(result.cl_granule_gr * scale);
    result.cl_tablete    = Math.round(result.cl_tablete * scale);
    result.ph_kg         = Math.round(result.ph_kg * scale * 100) / 100;
    result.antialgic_l   = Math.round(result.antialgic_l * scale * 100) / 100;
    result._extrapolated = true;
  }
  // Flag if pH was clamped
  if (ph !== phC) {
    result._phClamped = true;
    result._extrapolated = true;
  }

  return result;
}

/**
 * Get parameter status badge for a water chemistry value.
 * @param {string} param - 'fac'|'ph'|'ta'|'ch'|'cya'|'cc'
 * @param {number} value
 * @returns {{ status: 'ok'|'warn'|'danger', label: string } | null}
 */
function getParameterStatus(param, value) {
  if (value === null || value === undefined || isNaN(value)) return null;
  const ranges = {
    fac: { ok: [1.0, 3.0], warn: [0.3, 5.0] },
    ph:  { ok: [7.2, 7.6], warn: [7.0, 7.8] },
    ta:  { ok: [80, 120],  warn: [60, 150]  },
    ch:  { ok: [200, 400], warn: [150, 500] },
    cya: { ok: [30, 50],   warn: [0.1, 80]  },
    cc:  { ok: [0, 0.2],   warn: [0, 0.5]   }
  };
  const r = ranges[param];
  if (!r) return null;
  if (value >= r.ok[0] && value <= r.ok[1]) return { status: 'ok',     label: 'OK'      };
  if (value >= r.warn[0] && value <= r.warn[1]) return { status: 'warn', label: 'Atenție' };
  return { status: 'danger', label: 'Pericol' };
}

/**
 * Get chlorine efficiency % at given pH (from guide table).
 * @param {number} ph
 * @returns {number} efficiency percent (0-100)
 */
function getPhEfficiency(ph) {
  if (ph <= 7.0) return 78;
  if (ph <= 7.2) return 68;
  if (ph <= 7.4) return 55;
  if (ph <= 7.6) return 45;
  if (ph <= 7.8) return 33;
  if (ph <= 8.0) return 21;
  return 10;
}
