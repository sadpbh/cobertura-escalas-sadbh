# sistema_cobertura_escala

Um sistema web completo chamado "Cobertura de Escalas — SAD BH" para gestão de escalas de profissionais de saúde do Serviço de Atenção Domiciliar de Belo Horizonte (SAD BH / Melhor em Casa).

CONTEXTO DO NEGÓCIO
O SAD BH possui 21 equipes multidisciplinares distribuídas pela cidade. Cada equipe tem profissionais fixos (médico, enfermeiro e técnicos de enfermagem. Medicos e enfermeiros trabalham 6 horas por dia e técnicos em escala de plantão). Quando um profissional se ausenta (férias, licença médica, vacância, etc), um "ferista" é escalado para cobrir. Há 3 médicos feristas e 5 enfermeiros feristas disponíveis. Técnicos de enfermagem têm gestão própria pelo enfermeiro de cada equipe — o gestor só recebe licenças médicas e autoriza extras quando necessário.

IDENTIDADE VISUAL

Cores principais extraídas do logo "Melhor em Casa": verde escuro #2e7d4f, laranja #d95f1a, azul #1a5fa6
Fundo cinza claro #f4f6f9, cards brancos
Fonte: Plus Jakarta Sans
Estilo: sistema de saúde moderno, limpo, profissional
Topbar verde escuro com logo e navegação
Sem dark mode


PERFIS DE USUÁRIO
Gestor — acesso total. Login via Google OAuth. Vê todos os módulos.
Ferista — acesso via Google OAuth. Vê apenas sua própria escala destacada na aba Feristas.
Equipe — acesso via Google OAuth. Vê apenas sua própria equipe destacada na aba Equipes.
Usuários não cadastrados veem tela de "Acesso negado".

ESTRUTURA DE DADOS
Equipes (21 no total)
Cada equipe tem nome e dois turnos (manhã e tarde), podendo ter ausências em cada turno independentemente. As equipes são: Barreiro 1, Barreiro 2, Barreiro 3 Paliativos, Centro Sul 1, Centro Sul 2, HOB 1, HOB 2 Prolongados, Leste 1, Leste 2, Nordeste 1, Nordeste 3, Noroeste, Norte 1, Norte 2, Oeste 1, Oeste 2, Pampulha 1, Pampulha 2, Pediátrica, Venda Nova 2, SAD BH.
Feristas
Campos: Nome, Categoria (medico/enfermeiro), Disponível (SIM/NAO), Conselho (CRM/COREN), Disponibilidade fixa semanal (grade seg-sex × manhã-tarde, cada célula SIM/NAO).
Ausências
Campos: Data, Equipe, Categoria (medico/enfermeiro/tecnico), Turno (MANHA/TARDE), Profissional ausente, Motivo (Férias/Licença médica/Licença maternidade/Vacância/Afastamento INSS/Outro), Observações.
Alocações (escala gerada)
Campos: Data, Turno, Equipe, Categoria, Ferista alocado, Tipo (COBERTURA/APOIO), Peso da equipe.
Extras autorizados
Campos: Data, Equipe, Turno, Categoria, Tipo de extra, Profissional, Motivo, Autorizado em.
Acessos
Campos: Email, Perfil (gestor/ferista/equipe), Nome, Referência (nome do ferista ou equipe correspondente).

MÓDULOS DO SISTEMA
1. PAINEL PÚBLICO (/painel)
Acessível a todos os perfis autenticados. Topbar com navegação de semana (seletor de semana nativo + setas). Quatro abas:
Aba Equipes

Grade semanal (seg-sex) com todas as 21 equipes
Cada equipe mostra manhã e tarde separadamente
Cada slot mostra: tipo (Cobertura/Apoio/Sem cobertura/Extra/Sem ausência), nome do ferista, categoria (médico/enfermeiro com pill colorida)
Um slot pode ter múltiplas entradas quando há ausências simultâneas de categorias diferentes
Cores: verde claro = cobertura, laranja claro = apoio, vermelho claro = sem cobertura, roxo claro = extra autorizado, cinza = sem ausência
Alerta amarelo no topo listando equipes sem cobertura
Filtros: busca por nome, manhã/tarde, status
Perfil Equipe vê sua equipe destacada

Aba Feristas

Cards por ferista com visão mensal, destacando a semana atual (slots de alocação)
Cada slot mostra: tipo (Cobertura/Apoio/Extra), equipe, turno
Dias sem alocação mostram "Sem alocação"
Semana atual destacada
Filtros: busca por nome, categoria
Perfil Ferista vê seu card no topo destacado

Aba Técnicos

Só mostra solicitações de extra de técnicos e licenças registradas
Não mostra escala completa (gerida pelo enfermeiro de cada equipe)

Aba Gestor (só perfil gestor)

Cards de métricas: coberturas, sem cobertura, extras autorizados, taxa de cobertura
Dashboard com as regionais de belo horizonte que altera a cor da regional de acordo com a cobertura
Resumo por categoria (médicos e enfermeiros)
Lista de ausencias não cobertas com botão "Autorizar extra"
Carrega automaticamente ao clicar na aba

2. REGISTRO DE AUSÊNCIAS (/ausencias)
Só gestor. Duas abas:
Nova ausência

Categoria: Médico / Enfermeiro / Técnico de enf. (muda campos disponíveis)
Equipe/Base (select com lista completa)
Turno (oculto para técnico)
Profissional ausente (texto livre)
Motivo (select com opções)
Tipo de período: contínuo (data início → fim, pula fins de semana para médico/enfermeiro) ou dias avulsos
Pills clicáveis mostrando dias que serão registrados (clique remove o dia)
Resumo ao lado mostrando o que será gravado
Botão salvar habilitado só quando todos os campos obrigatórios preenchidos

Ausências registradas

Tabela com todas as ausências
Filtros por categoria e motivo
Botão excluir por linha
Cards de métricas no topo (total, por categoria)

3. GESTÃO DE ESCALA (/escala)
Só gestor. Três etapas em stepper:
Etapa 1 — Gerar

Seletor de data (segunda-feira da semana)
Botão "Gerar escala" que chama o motor de alocação
Log de execução em tempo real
Botão "Ver escala atual"

Etapa 2 — Revisar e ajustar

Cards de métricas (coberturas, sem cobertura, apoios, taxa)
Lista de buracos não cobertos
Grade de alocações: equipes × dias, com slots clicáveis
Clicar num slot abre modal para trocar ferista
Modal mostra feristas disponíveis (com indicação de ocupação no dia) e opção "Remover cobertura"
Badge "editado" nos slots alterados
Contador de alterações pendentes
Botão "Salvar ajustes"

Etapa 3 — Publicar

Resumo final com métricas
Alerta se houver buracos não cobertos ou ajustes não salvos
Botão "Publicar escala"
Tela de sucesso após publicação

4. GERENCIAR FERISTAS (/feristas)
Só gestor.
Formulário de cadastro/edição:

Nome completo, Categoria, Disponível no sistema (ativo/inativo), Número do conselho
Grade de disponibilidade fixa: seg-sex × manhã-tarde
Cada célula é um botão toggle verde (disponível) / cinza (indisponível)
Por padrão todos os turnos marcados como disponíveis

Tabela de feristas cadastrados:

Nome, Categoria (pill colorida), Conselho, Status (ativo/inativo)
Mini visualização da disponibilidade (10 pontos coloridos: 5 manhãs + 5 tardes)
Botões Editar e Excluir
Confirmação antes de excluir

5. GERENCIAR ACESSOS (/acessos)
Só gestor.

Formulário: email, perfil, nome, referência (equipe ou ferista, dependendo do perfil)
Tabela de usuários com botão remover
Descrição de cada perfil

6. MAPA MENSAL DE COBERTURA (dentro do Gestor)
Grade mensal contínua:

Linhas = equipes, Colunas = dias do mês (corrido, sem separação de semana)
Duas grades separadas: uma para médicos, uma para enfermeiros
Cores por célula: verde = coberto, laranja = apoio/extra, vermelho = sem cobertura, cinza = sem ausência
Semana atual destacada com borda
Navegação por mês
Atualização automática conforme ausências e escalas

7. VISÃO MENSAL DOS FERISTAS

Calendário mensal mostrando todos os dias
Semana atual destacada
Cada dia mostra o status de alocação do ferista
Contínuo entre meses (sem quebra de página)


LÓGICA DE ALOCAÇÃO (motor de escala)
O motor respeita esta ordem de prioridade:

Regras OBRIGATORIO — ferista deve ser alocado em equipe/turno específico
Regras NAO_ALOCAR — ferista bloqueado para dia/equipe/turno
Disponibilidade fixa do cadastro (grade semanal)
Regras PERMITIR_APENAS — ferista só pode atuar em turnos específicos
Regra DOBRAR — permite ferista atuar em dois turnos no mesmo dia

Critérios de scoring para escolha do melhor ferista:

Equipes com maior peso têm prioridade
Ferista com menos alocações na semana tem prioridade
Continuidade (mesmo ferista na mesma equipe em dias consecutivos tem bônus)
Ferista não deve repetir na mesma equipe mais de 2 vezes na semana se houver alternativa
Um ferista pode atuar em dois turnos no mesmo dia apenas se tiver regra DOBRAR


REGRAS DE NEGÓCIO IMPORTANTES

Médicos e enfermeiros trabalham seg-sex
Técnicos trabalham em esquema de plantão (inclui fins de semana)
Ausências podem ser programadas (férias) ou emergenciais (licença médica de última hora)
A escala é gerada semanalmente mas sofre ajustes durante a semana
Quando não há ferista disponível, o gestor pode autorizar um extra (médico, enfermeiro ou técnico)
O gestor é o único que registra ausências e gera escalas
Feristas e equipes consultam mas não editam


COMPONENTES DE UI NECESSÁRIOS

Topbar verde com logo, navegação de semana e menu do gestor
Stepper de 3 etapas para geração de escala
Grade semanal responsiva com slots clicáveis
Modal de troca de ferista com lista de disponibilidade
Pills de dias (clicáveis para remover)
Toggle grid para disponibilidade de feristas
Cards de métricas com indicadores coloridos
Log de execução em tempo real
Alert bars coloridas (orange para avisos, red para erros, green para sucesso)
Mini-indicadores de disponibilidade (10 pontos)
Seletor de semana com setas de navegação
Grade mensal de cobertura com gradiente de cores
Tela de acesso negado


OBSERVAÇÕES TÉCNICAS

Backend via Google Apps Script (API REST já construída)
Comunicação via google.script.run para chamadas internas
Autenticação via Google OAuth (Session.getActiveUser().getEmail())
Dados armazenados em Google Sheets (abas: cad_equipes, cad_feristas, escala_base, ALOCACAO_AUTO, cad_extras, ausencias_tecnicos, cad_acessos, REGRAS_SEMANA)
Sistema já em produção parcial — o frontend deve ser compatível com a API existente
Todos os textos em português brasileiro
Sem dark mode
Responsivo mas otimizado para desktop
