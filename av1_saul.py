import gspread
import pandas as pd

filename = "service_account.json"

df = pd.DataFrame({'nomes': ['João Silva silva', 'Maria Santos', 'José Silva', 'Pedro Alves']})

# dividir a coluna 'nomes' em duas colunas usando um espaço como separador
df[['primeiro_nome', 'sobrenome', 'sobrenome2']] = df['nomes'].str.split(' ', expand=True)

df[['mes','dia','ano']] = df['data de nascimento'].str.split('/', expand=True)

# exibir o resultado
print(df)