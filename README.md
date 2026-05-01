# Sistema de Pericias - Integra Contabilidade e Pericias

Painel web para monitoramento de intimacoes judiciais (TJES/TRE-ES) e geracao de laudos periciais contabeis via Claude AI.

## Arquivos

- painel/index.html - Painel da equipe (qualquer navegador)
- - painel/cowork.html - Painel do perito (Claude Cowork)
  - - dados/pericias.json - Dados atualizados pelo agente Claude
    - - agente/processar_pericias.py - Script cruzamento Gmail x Excel
      - - servidor/setup.sh - Instalacao automatica (Ubuntu 22.04)
       
        - ## Servidor
       
        - - Hetzner CX22 - Ubuntu 22.04, Frankfurt
          - - Caddy - HTTPS automatico (Let's Encrypt)
            - - URL: https://pericias.integraconsult.com.br
             
              - ## Deploy
             
              - git add . && git commit -m "update" && git push
             
              - Desenvolvido com Claude AI - Integra Contabilidade & Pericias - 2026
