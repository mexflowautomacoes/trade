# Proposta — iMoneyWeb

---

Fala Delmer, beleza? Montei a proposta do projeto novo com tudo que conversamos. Lê com calma e me chama pra tirar qualquer dúvida.

---

## O que é o iMoneyWeb

O iMoneyWeb é a evolução do FlowTrader — mesma essência de dashboard de fluxo de agressão em tempo real, mas agora rodando na nuvem, conectado direto na API da CEDRO.

Você acessa de qualquer lugar: computador, celular, tablet. É só abrir o link.

Sem instalar nada. Sem depender do Excel aberto. Sem ficar preso a uma máquina.

---

## O que vai ter

**Dados em tempo real**
- Conexão direta com a API da CEDRO via Socket — dados de mercado ao vivo, sem delay
- Reconexão automática se a conexão cair

**Multi-ativo**
- Escolha do ativo por código (mini índice, mini dólar, ações, etc.)
- Troca de ativo ao vivo, sem reiniciar
- Até 2 ativos simultâneos na tela
- Opção de sobrepor os 2 gráficos na mesma tela ou ver lado a lado

**Filtros inteligentes**
- Filtro por quantidade mínima — só contabiliza agressões acima de X contratos
- Filtro por corretora — seleciona quais corretoras você quer acompanhar (ex: só XP, UBS e JP Morgan). O sistema filtra e só processa os negócios dessas corretoras

**Dois modos de contagem**
- Modo sinal: cada agressão vale +1 (compra) ou -1 (venda), gerando o saldo direcional
- Modo volume: contabiliza a quantidade real de contratos. Veio uma ordem de 150? Soma 150. Veio outra de 178? Mais 178. Assim você vê, por exemplo, que a UBS está com 24 mil contratos comprados no período

**Dashboard completo**
- Gráfico de fluxo com o somatório ao longo do pregão
- KPIs em tempo real (saldo, volume, quantidade, etc.)
- Tape de negócios (últimos trades)
- Ranking de corretoras com volume acumulado — quem está comprando mais, quem está vendendo mais

**Acesso e segurança**
- Login com senha pra proteger seu acesso
- Responsivo — funciona no PC, tablet e celular
- Online 24h, hospedado em servidor na nuvem

---

## Por que o iMoneyWeb é diferente do FlowTrader

Delmer, o FlowTrader é um script que roda local na sua máquina, lendo dados do Excel que já está aberto. Funciona bem pro que foi feito, mas é limitado ao seu computador.

O iMoneyWeb é outro nível de produto. Vou te explicar o que está envolvido:

**Integração com a API da CEDRO** — O sistema conecta direto na API profissional de market data da CEDRO via Socket, recebendo os negócios em tempo real. Isso exige desenvolver toda a camada de conexão, autenticação, tratamento dos dados e reconexão automática caso caia. Não é só ler uma planilha — é manter uma conexão de streaming aberta processando dados ao vivo.

**Servidor na nuvem** — O FlowTrader roda no seu PC. O iMoneyWeb vai rodar num servidor dedicado na internet, online 24 horas. Preciso contratar, configurar e manter esse servidor: sistema operacional, segurança, certificado SSL, domínio, deploy da aplicação.

**Acesso de qualquer dispositivo** — O dashboard precisa funcionar no computador, no celular e no tablet. Isso significa desenvolver uma interface responsiva que se adapta a cada tela — é um trabalho considerável em relação ao FlowTrader que só roda no PC.

**Multi-ativo com filtros avançados** — O FlowTrader trabalha com 1 ativo e filtros básicos. Aqui são 2 simultâneos com sobreposição de gráficos, filtro por corretora com seleção múltipla, modo volume com acumulado real por corretora. Toda a lógica de dados, gráficos e filtros precisa funcionar em paralelo.

**Autenticação** — Como fica na internet, precisa de login e senha pra garantir que só você acessa.

---

## O que já vem incluso

Além do desenvolvimento da plataforma, o iMoneyWeb já inclui toda a infraestrutura necessária pra funcionar:

- **Servidor dedicado na nuvem** — máquina online 24h que roda a aplicação, já contratada e configurada por mim
- **Domínio próprio** — você acessa pelo seu endereço exclusivo (ex: imoneyweb.com.br ou o domínio que preferir)
- **Certificado SSL (HTTPS)** — conexão segura e criptografada, aquele cadeado verde no navegador
- **Deploy e configuração** — toda a parte de colocar o sistema no ar: instalação, configuração do servidor, firewall, atualizações de segurança
- **Monitoramento** — se o servidor cair ou a aplicação travar, eu fico sabendo e resolvo
- **Backups** — configuração de backup automático pra não perder dados
- **Manutenção técnica contínua** — atualizações, correções, ajustes na API da CEDRO se ela mudar
- **Suporte via WhatsApp** — qualquer problema ou dúvida, é só me chamar

Você não precisa se preocupar com nada técnico. Eu cuido de tudo.

---

## Investimento

Preparei dois planos pra você escolher o que faz mais sentido. Nos dois, o resultado final é o mesmo — o iMoneyWeb completo, com tudo que descrevi acima.

---

### Plano 1 — Mensal completo

Tudo incluso num valor mensal. Sem taxa de entrada.

| | |
|---|---|
| **Valor** | **R$ 997/mês** |
| **Compromisso** | 12 meses |
| **Início da cobrança** | No início do projeto |

Inclui: desenvolvimento, servidor, domínio, SSL, manutenção e suporte.

**Por que esse plano é interessante:** Você não tem nenhum valor de entrada. É como assinar uma ferramenta de trabalho — R$ 997/mês e pronto. Se você usa o iMoneyWeb pra operar melhor, esse valor se paga num dia bom de pregão.

---

### Plano 2 — Entrada + mensal reduzido

Uma entrada que cobre o desenvolvimento, e um mensal mais leve que cobre toda a infraestrutura e manutenção.

| | |
|---|---|
| **Entrada** | **R$ 3.500** (em 2x de R$ 1.750) |
| **Mensal** | **R$ 700/mês** (a partir da entrega) |
| **Compromisso mensal** | 12 meses após a entrega |

A entrada cobre o desenvolvimento. O mensal cobre: servidor, domínio, SSL, manutenção e suporte.

**Por que esse plano é interessante:** Você paga menos por mês depois da entrega. A entrada cobre o desenvolvimento e o mensal fica mais leve no fluxo de caixa.

---

### API CEDRO — por sua conta

Você contrata direto com a CEDRO o plano de market data. Antes de fechar, vamos confirmar juntos que o plano escolhido cobre o streaming de negócios com identificação de corretoras — que é o dado essencial pro iMoneyWeb funcionar.

- Plano Basic anual: ~R$ 440/mês
- Plano Basic mensal: ~R$ 660/mês
- Site: marketdatacloud.com.br

---

## Uma conta que vale fazer

Delmer, o Profit Chart Pro custa em torno de R$ 350/mês. Se hoje você usa o Profit principalmente pra acompanhar fluxo de agressão, com o iMoneyWeb rodando na nuvem — com filtro por corretora, modo volume, ranking acumulado, 2 ativos simultâneos — vale avaliar se o Profit ainda faz sentido pra você.

Se decidir cancelar, essa economia já cobre boa parte da mensalidade do iMoneyWeb. No Plano 1, por exemplo, o custo real cai de R$ 997 pra algo em torno de R$ 647/mês.

A decisão é sua — mas a conta fecha.

---

## Condições importantes

- **Disponibilidade da CEDRO:** indisponibilidades, mudanças radicais ou descontinuação da API CEDRO estão fora do meu controle. Se a CEDRO mudar o formato dos dados, eu ajusto (isso entra na manutenção). Se a CEDRO ficar fora do ar, a responsabilidade é deles.
- **Cancelamento no Plano 1:** o compromisso é de 12 meses. Cancelamento antecipado não isenta das parcelas restantes.
- **Cancelamento no Plano 2:** o valor de entrada não é reembolsável. O mensal tem compromisso de 12 meses após a entrega.
- **Após os 12 meses:** o serviço continua mês a mês. Você pode cancelar a qualquer momento com 30 dias de aviso.

---

Qualquer dúvida, me chama. Se quiser, a gente marca uma call rápida pra alinhar.

Abraço!
