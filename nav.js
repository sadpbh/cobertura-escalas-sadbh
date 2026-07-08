// ============================================================
// Cobertura de Escalas · SAD-BH
// nav.js — menu de navegação compartilhado entre as telas.
// Chamar renderNav('acessos'|'pessoas'|'ausencias'|'escala')
// depois que window.meuAcesso já estiver resolvido (dentro de
// authReady).
// ============================================================
function renderNav(paginaAtiva) {
  var el = document.getElementById('nav-menu');
  if (!el) return;

  var perfil = window.meuAcesso ? window.meuAcesso.perfil : null;

  var itensPorPerfil = {
    gestor: [
      { chave: 'painel', label: 'Painel', href: 'painel.html' },
      { chave: 'equipes', label: 'Equipes', href: 'equipes.html' },
      { chave: 'acessos', label: 'Acessos', href: 'acessos.html' },
      { chave: 'pessoas', label: 'Pessoas', href: 'pessoas.html' },
      { chave: 'ausencias', label: 'Ausências', href: 'ausencias.html' },
      { chave: 'escala', label: 'Escala', href: 'escala.html' },
      { chave: 'ocorrencias', label: 'Ocorrências', href: 'ocorrencias.html' }
    ],
    equipe: [
      { chave: 'ausencias', label: 'Ausências', href: 'ausencias.html' },
      { chave: 'escala', label: 'Escala', href: 'escala.html' }
    ],
    pessoa: [
      { chave: 'minha_escala', label: 'Minha Escala', href: 'minha_escala.html' }
    ]
  };

  var itens = itensPorPerfil[perfil] || [];
  if (itens.length === 0) { el.innerHTML = ''; return; }

  el.innerHTML = itens.map(function (item) {
    var ativo = item.chave === paginaAtiva;
    return '<a href="' + item.href + '" class="nav-link' + (ativo ? ' nav-link-ativo' : '') + '">' + item.label + '</a>';
  }).join('');
}
