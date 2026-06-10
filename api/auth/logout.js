import { clearSession } from '../_utils.js';

export default function handler(req, res) {
  clearSession(res);
  console.log('[Auth] User logged out');
  res.redirect('/');
}
