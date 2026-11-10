// Endereço e chave pública (publishable) do Supabase.
// Versão Vercel: aponta para o MESMO projeto Supabase do Perfil de Sondagem
// (mesmo login, mesma empresa). A chave publishable é pública por natureza —
// quem protege os dados de cada empresa são as regras de segurança (RLS) do banco.
// A configuração antiga (servidor soiltec, ambiente dev) ficou em config.soiltec-dev.js.
window.APP_CONFIG = {
  supabaseUrl: 'https://wbocmpekamjfwkwmxtjd.supabase.co',
  supabaseAnonKey: 'sb_publishable_rackRbyFgs42jUbWBu0Krw_tn7VTJOw'
};
