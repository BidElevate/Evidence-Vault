import { createClient } from '@supabase/supabase-js';

const SUPABASE_URL = 'https://oyagzfirgdlgmihssuyr.supabase.co';
const SUPABASE_KEY = 'sb_publishable_sErw8WmyBQtjaD1XSiaw9w_1vRp7WYw';

export const supabase = createClient(SUPABASE_URL, SUPABASE_KEY);
