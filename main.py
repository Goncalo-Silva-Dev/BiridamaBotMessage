import pandas as pd
from instagrapi import Client

# Configurações de login no Instagram
USERNAME = "testerbot2025"
PASSWORD = "tester_bot*2025"

# Caminho do arquivo Excel
file_path = "pedidosJourneyETBvteste.xlsx"

# Carregar todas as folhas do Excel
sheets = pd.read_excel(file_path, sheet_name=None)

# Inicializar cliente Instagram
cl = Client()
try:
    cl.login(USERNAME, PASSWORD)
except Exception as e:
    print(f"❌ Erro ao fazer login no Instagram: {e}")
    exit()

# Função para obter detalhes dos itens do pedido
def obter_itens_pedido(row):
    itens = []
    for i in range(1, 5):  # Assumindo até 4 produtos por pedido
        produto = row.get(f"Produto{i}", "").strip() if isinstance(row.get(f"Produto{i}"), str) else ""
        quantidade = row.get(f"Quantidade{i}", 0)
        preco = row.get(f"Preço unitario{i}", 0)
        
        if pd.notna(produto) and produto and pd.notna(quantidade) and pd.notna(preco):
            try:
                quantidade = int(quantidade)
                preco = float(preco)
                total_item = quantidade * preco
                itens.append(f"{quantidade}x {produto} - {preco:.2f}€ cada (Total: {total_item:.2f}€)")
            except ValueError:
                print(f"⚠️ Erro ao processar produto {produto}: {quantidade}, {preco}")
                continue
    return "\n".join(itens) if itens else "Nenhum produto encontrado."

# Função para criar mensagem baseada no tipo selecionado no Excel
def criar_mensagem(row):
    tipo_mensagem = str(row.get("Tipo de Mensagem", "")).strip().lower()  # Garantido que seja uma string
    
    # Verificar se "Envio ?" é 1 para definir o custo de envio
    envio_custo = row.get("Envio €", 0) if row.get("Envio ?", 0) == 1 else 0
    
    pagamento_status = "Pago" if row.get("Pago ?", 0) == 1 else "Por Pagar"
    itens_pedido = obter_itens_pedido(row)
    
    total = row.get("Total")
    
    if pd.isna(total) or total is None:
        total = sum([int(row.get(f"Quantidade{i}", 0)) * float(row.get(f"Preço unitario{i}", 0)) 
                    for i in range(1, 5) 
                    if pd.notna(row.get(f"Produto{i}")) and row.get(f"Produto{i}") != ""])
    
    if total is None or not isinstance(total, (int, float)):
        total = 0.0

    mensagens = {
        "nova_reserva": f"Olá {row['User']}, a sua reserva foi registada.\n\n{itens_pedido}\n\nValor total dos artigos: {total - envio_custo:.2f}€\nEnvio: {envio_custo:.2f}€\nEstado de pagamento: {pagamento_status}\nTotal: {total:.2f}€\nObrigado pela sua compra!",
        "reserva_alterada": f"Olá {row['User']}, a sua reserva foi alterada.\n\n{itens_pedido}\n\nValor total dos artigos: {total - envio_custo:.2f}€\nEnvio: {envio_custo:.2f}€\nEstado de pagamento: {pagamento_status}\nTotal: {total:.2f}€\nObrigado pela sua compra!",
        "pagamento_recebido": f"Olá {row['User']}, recebemos o pagamento da sua reserva.\n\n{itens_pedido}\n\nValor total dos artigos: {total - envio_custo:.2f}€\nEnvio: {envio_custo:.2f}€\nEstado de pagamento: Pago\nTotal: {total:.2f}€\nObrigado pela sua compra!",
        "aviso_pagamento": f"Olá {row['User']}, lembramos que o pagamento da sua reserva ainda não foi recebido.\n\n{itens_pedido}\n\nValor total dos artigos: {total - envio_custo:.2f}€\nEnvio: {envio_custo:.2f}€\nEstado de pagamento: Por Pagar\nTotal: {total:.2f}€\nAgradecemos a sua atenção!",
    }

    return mensagens.get(tipo_mensagem, "Tipo de mensagem não reconhecido.")

# Loop para processar cada linha do DataFrame
for sheet_name, df in sheets.items():
    for index, row in df.iterrows():
        if row.get("Enviado", 0) == 0:  # Verifica se a mensagem ainda não foi enviada
            mensagem = criar_mensagem(row)
            try:
                username = row['User'].lstrip("@").strip()  # Remove '@' se presente
                try:
                    user_id = cl.user_id_from_username(username)  # Obtém o ID do usuário
                except Exception:
                    print(f"⚠️ Usuário {row['User']} não encontrado. Pulando...")
                    continue
                
                # Enviar a mensagem para o ID do usuário
                cl.direct_send(mensagem, [user_id])  
                
                # Marca como enviado
                df.at[index, "Enviado"] = 1  # Atualiza a coluna "Enviado" para 1 após o envio
                print(f"✅ Mensagem enviada para {row['User']}")
                
            except Exception as e:
                print(f"⚠️ Erro ao enviar mensagem para {row['User']}: {e}")
                df.at[index, "Enviado"] = 1  # Atualiza a coluna "Enviado" para 1 após o envio

# Salvar todas as folhas modificadas no arquivo Excel
with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
    for sheet_name, df in sheets.items():
        df.to_excel(writer, sheet_name=sheet_name, index=False)  # Escrever a folha no arquivo

# Após finalizar todos os processos, o cliente faz logout
cl.logout()
print("✅ Atualizações salvas no arquivo Excel.")
