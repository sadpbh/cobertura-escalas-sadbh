// ============================================================
// Cobertura de Escalas · SAD-BH
// auth.js — Login por e-mail + senha (Supabase Auth) + resolução
// de perfil via RPC meu_acesso().
// Requer supabase-client.js carregado antes.
//
// Observação operacional: os usuários NÃO se autocadastram.
// O gestor cria cada login em Authentication > Users > Add user
// no painel do Supabase, definindo a senha ali diretamente —
// isso evita qualquer dependência de envio de e-mail no login
// do dia a dia. Redefinição de senha esquecida também é feita
// manualmente pelo gestor no mesmo painel.
// ============================================================
(function () {
  var _pending = [];
  var _done = false;

  // ---- API pública ----
  window.authReady = function (fn) {
    if (_done) { fn(); } else { _pending.push(fn); }
  };

  window.meuAcesso = null; // { perfil, equipe_id, equipe_nome, pessoa_id, pessoa_nome }

  window.authLogout = function () {
    supabaseClient.auth.signOut().then(function () {
      window.location.reload();
    });
  };

  // ---- Internos ----
  function _fire() {
    _done = true;
    var cbs = _pending.slice(); _pending = [];
    cbs.forEach(function (fn) { fn(); });
  }

  function _el(id) { return document.getElementById(id); }

  function _injectOverlay() {
    if (_el('auth-overlay')) return;
    var div = document.createElement('div');
    div.id = 'auth-overlay';
    div.innerHTML =
      '<div class="auth-card">' +
        '<div class="auth-icon">&#x1F3E5;</div>' +
        '<h2>Cobertura de Escalas</h2>' +
        '<p class="auth-sub">SAD BH</p>' +
        '<div id="auth-step-login">' +
          '<label class="auth-label">E-mail</label>' +
          '<input id="auth-email-input" type="email" placeholder="email@pbh.gov.br" autocomplete="username">' +
          '<label class="auth-label">Senha</label>' +
          '<input id="auth-password-input" type="password" placeholder="Sua senha" autocomplete="current-password">' +
          '<button id="auth-btn-entrar">Entrar</button>' +
        '</div>' +
        '<div id="auth-step-negado" style="display:none">' +
          '<p class="auth-negado-msg">Seu e-mail não está autorizado a acessar este sistema. Fale com o gestor do SAD BH.</p>' +
          '<button id="auth-btn-sair">Sair</button>' +
        '</div>' +
        '<p id="auth-status"></p>' +
      '</div>';
    document.body.insertBefore(div, document.body.firstChild);

    _el('auth-btn-entrar').addEventListener('click', _entrar);
    _el('auth-password-input').addEventListener('keydown', function (e) {
      if (e.key === 'Enter') _entrar();
    });
  }

  function _setStatus(msg, isError) {
    var el = _el('auth-status');
    if (!el) return;
    el.textContent = msg || '';
    el.style.color = isError ? '#c0392b' : '#6b7a96';
  }

  function _entrar() {
    var email = (_el('auth-email-input').value || '').trim();
    var senha = _el('auth-password-input').value || '';
    if (!email || !senha) { _setStatus('Informe e-mail e senha.', true); return; }
    _setStatus('Entrando...', false);
    _el('auth-btn-entrar').disabled = true;

    supabaseClient.auth.signInWithPassword({ email: email, password: senha }).then(function (res) {
      _el('auth-btn-entrar').disabled = false;
      if (res.error) {
        _setStatus('E-mail ou senha inválidos.', true);
        return;
      }
      _resolverAcesso();
    });
  }

  function _mostrarNegado() {
    _el('auth-step-login').style.display = 'none';
    _el('auth-step-negado').style.display = 'block';
    _el('auth-btn-sair').addEventListener('click', window.authLogout);
  }

  function _resolverAcesso() {
    return supabaseClient.rpc('meu_acesso').then(function (res) {
      if (res.error || !res.data || res.data.length === 0) {
        _mostrarNegado();
        return false;
      }
      window.meuAcesso = res.data[0];
      var overlay = _el('auth-overlay');
      if (overlay) overlay.remove();
      _fire();
      return true;
    });
  }

  function _init() {
    _injectOverlay();
    supabaseClient.auth.getSession().then(function (res) {
      var session = res.data && res.data.session;
      if (session) {
        _resolverAcesso();
      }
      // se não há sessão, o formulário de login (já injetado) permanece visível
    });
  }

  document.addEventListener('DOMContentLoaded', _init);
})();

