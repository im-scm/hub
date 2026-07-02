# SCM Analytics Hub - Clean Final

Esta é a versão revisada para evitar navegação duplicada.

## Estrutura correta

- app.py
- requirements.txt
- Cockpit_Papel.xlsm
- app_paperbase.xlsx
- assets/impress_logo.png
- pages/01_Cockpit_Papel.py
- pages/02_Paper_Base.py

## Importante

- O app.py é o único arquivo que define `st.navigation`.
- As páginas não usam `st.page_link`, `st.switch_page` nem `st.set_page_config`.
- Cada página contém apenas o dashboard e seus filtros.
