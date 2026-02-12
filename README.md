# Dashboard de Carga Máquina (Simulação de Cenários)

Este projeto lê um **XLSX de carga máquina** e gera um **dashboard em Streamlit** com:

- Filtros pelas colunas **C, F, J, R** (inclui **Modelo pela coluna C**)
- Parâmetros do cenário: **OEE / Eficiência Máquina** e **Eficiência de Mão de Obra**
- Capacidade por **turnos (horas)** e **dias úteis**
- Gráficos em padrão “indústria”:
  - **Carga (horas trabalhadas)** por agrupamento (com linha de capacidade)
  - **Gráfico de barras somando o TAKT** (convertido para horas)

## Como rodar

```bash
python -m venv .venv
# Windows:
.venv\Scripts\activate

pip install -r requirements.txt
streamlit run app.py
```

Abra o link mostrado no terminal (geralmente http://localhost:8501).

## Observações

- O cálculo de **horas trabalhadas** usa: `QTD TOTAL MINUTOS / 60`.
- A **capacidade efetiva** é: `(horas turnos × dias úteis) × OEE × Eficiência MO`.
