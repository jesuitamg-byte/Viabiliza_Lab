# Política de segurança

## Versões com suporte

Só a versão mais recente recebe correções, inclusive de segurança. Quem usa uma versão antiga deve atualizar: na versão on-line basta recarregar a página; no Windows, basta instalar o instalador novo por cima (os projetos são mantidos).

| Versão  | Suporte            |
| ------- | ------------------ |
| 9.0.x   | :white_check_mark: |
| < 9.0   | :x:                |

## Como o programa trata os dados

- O ViabilizaLab não tem servidor próprio nem conta de usuário. A versão on-line roda inteira no navegador de quem a usa, e a versão instalada roda no próprio computador.
- Os projetos ficam com o usuário: arquivos `.pnbox` em Documentos\ViabilizaLab, na versão instalada, ou no próprio navegador, na versão on-line.
- O programa consulta apenas serviços públicos e gratuitos (IBGE, Banco Central e OpenStreetMap) e envia só o necessário para a consulta, como o código do município. Nenhum dado do projeto é enviado.
- O texto preparado para análise com IA só sai do programa se o próprio usuário o copiar e colar em outro serviço.

## Como relatar uma vulnerabilidade

**Não abra uma issue pública** para relatar falhas de segurança.

Use o relato privado do GitHub: na aba **Security** deste repositório, clique em **Report a vulnerability**. Descreva:
- o que acontece e como reproduzir o problema;
- a versão do programa (aparece no menu lateral) e se é a versão on-line ou a instalada;
- o impacto que você imagina (por exemplo, acesso a arquivos ou execução de código).

## O que esperar

- Confirmação de recebimento em até **7 dias**.
- Uma primeira avaliação em até **30 dias**, dizendo se o problema foi aceito ou não e o motivo.
- Problemas aceitos são corrigidos numa nova versão, registrada no [histórico de versões](CHANGELOG.md). Quem relatou recebe o crédito, se quiser.

O ViabilizaLab é um projeto acadêmico mantido sem fins lucrativos; os prazos acima são compromissos de boa-fé.

## Fora do escopo

- O aviso "O Windows protegeu o computador" ao instalar: ele aparece porque o instalador não tem assinatura digital paga, e o programa não usa esse recurso pago.
- Falhas nos serviços externos consultados (IBGE, Banco Central, OpenStreetMap), que devem ser relatadas aos responsáveis por eles.
