import pandas as pd
import psycopg2
import schedule
import time
from datetime import datetime
import openpyxl

# Configurações de conexão com o banco de dados
DB_CONFIG = {
    'host': 'seu_host',
    'database': 'seu_banco',
    'user': 'seu_usuario',
    'password': 'sua_senha'
}
# Armazena os IDs da verificação anterior
ids_anteriores = set()

def verificar_tabela():
    try:
        # Estabelece conexão com o banco
        conn = psycopg2.connect(**DB_CONFIG)
        cursor = conn.cursor()
        
        # Substitua 'sua_tabela' e 'coluna_id' pelos nomes reais
        query = "SELECT * FROM sua_tabela"
        
        # Lê a tabela do banco para um DataFrame
        df = pd.read_sql_query(query, conn)
        
        # Obtém os IDs atuais
        ids_atuais = set(df['coluna_id'].values)
        
        # Encontra novos IDs
        novos_ids = ids_atuais - ids_anteriores
        
        # Marca as linhas novas
        df['é_novo'] = df['coluna_id'].isin(novos_ids)
        
        # Atualiza os IDs anteriores
        global ids_anteriores
        ids_anteriores = ids_atuais
        
        # Gera nome do arquivo com timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        nome_arquivo = f'tabela_atualizada_{timestamp}.xlsx'
        
        # Exporta para Excel, destacando as linhas novas
        with pd.ExcelWriter(nome_arquivo, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)
            
            # Formatação do Excel
            workbook = writer.book
            worksheet = writer.sheets['Sheet1']
            
            # Destaca linhas novas em amarelo
            for idx, row in df.iterrows():
                if row['é_novo']:
                    for col in range(len(df.columns)):
                        worksheet.cell(row=idx+2, column=col+1).fill = \
                            openpyxl.styles.PatternFill(start_color='FFFF00', 
                                                      end_color='FFFF00',
                                                      fill_type='solid')
        
        print(f"Arquivo {nome_arquivo} gerado com sucesso!")
        
        cursor.close()
        conn.close()
        
    except Exception as e:
        print(f"Erro: {str(e)}")

# Agenda a execução a cada 15 minutos
schedule.every(15).minutes.do(verificar_tabela)

# Loop principal
if __name__ == "__main__":
    print("Monitoramento iniciado...")
    verificar_tabela()  # Primeira execução
    while True:
        schedule.run_pending()
        time.sleep(1)