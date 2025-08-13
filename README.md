# Consulta-Usuario-AD
Consulta de Usuários e Maquinas dentro de um período estabelecido usando o Powershell

Como analista de rede atualmente monte um script pra otimizar meu tempo!

🚀 Script PowerShell para Relatórios de Usuários e Computadores no Active Directory
Desenvolvi um script em PowerShell que gera relatórios completos de usuários e computadores ativos/inativos no Active Directory, com base em uma data limite de logon informada pelo administrador.
🛠 O que ele faz:
Solicita uma data limite e valida se não é futura.
Lista usuários inativos no AD, com nome, login e data do último logon.
Lista computadores inativos, ativos e logados no dia.
Pode buscar todas as máquinas, filtrar por nome digitado ou usar uma lista de nomes de máquinas a partir de um arquivo TXT.
Exporta automaticamente os relatórios para arquivos .txt na área de trabalho.
Apresenta no console os resultados formatados e com contagem total por categoria.
📂 Saída dos relatórios:
Usuários Inativos 📴
Máquinas Inativas (sem logon desde a data limite)
Máquinas Ativas (logon entre a data limite e a data atual)
Máquinas Logadas Hoje
📌 Exemplo de uso:
Informe a data limite (ex.: 01/01/2024).
Informe um nome de máquina ou deixe em branco para listar todas.
Receba relatórios claros e organizados, prontos para auditorias e inventário.
