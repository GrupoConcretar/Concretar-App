// ============================================================
// CONFIGURACIÓN DE SUPABASE
// Reemplazá estos dos valores por los de tu proyecto:
// Project Settings > API > Project URL / anon public key
// ============================================================
export const SUPABASE_URL = "https://ijkthelpxsyjicanduaj.supabase.co";
export const SUPABASE_ANON_KEY = "sb_publishable_YTPsdNI7QefbSm6YkDiaaA_1NCjtXnV";

export const isSupabaseConfigured =
  SUPABASE_URL.startsWith("http") &&
  !SUPABASE_URL.includes("TU-PROYECTO") &&
  SUPABASE_ANON_KEY &&
  !SUPABASE_ANON_KEY.includes("TU_ANON_KEY");
