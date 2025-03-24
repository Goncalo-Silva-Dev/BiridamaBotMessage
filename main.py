import pandas as pd
import time
from instagrapi import Client

# Configurações de login no Instagram
USERNAME = "testerbot2025"
PASSWORD = "tester_bot*2025"

# Caminho do arquivo Excel
file_path = "pedidosJourneyETB.xlsx"

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
    tipo_mensagem = row.get("Tipo de Mensagem", "").strip().lower()
    
    # Verificar se "Envio ?" é 1 para definir o custo de envio
    envio_custo = row.get("Envio €", 0) if row.get("Envio ?", 0) == 1 else 0
    
    pagamento_status = "Pago" if row.get("Pago ?", 0) == 1 else "Por Pagar"
    itens_pedido = obter_itens_pedido(row)
    
    # Tente obter o total da coluna "Total", se não estiver disponível, calcule a partir dos itens
    total = row.get("Total")
    
    # Se o total não estiver disponível ou não for um número, calcule a partir dos itens
    if pd.isna(total) or total is None:
        total = sum([int(row.get(f"Quantidade{i}", 0)) * float(row.get(f"Preço unitario{i}", 0)) 
                    for i in range(1, 5) 
                    if pd.notna(row.get(f"Produto{i}")) and row.get(f"Produto{i}") != ""])
    
    # Garantir que total é um número
    if total is None or not isinstance(total, (int, float)):
        total = 0.0  # Defina um valor padrão se total não for válido

    mensagens = {
        "nova_reserva": f"Olá {row['User']}, a sua reserva foi registada.\n\n{itens_pedido}\n\nValor total dos artigos: {total:.2f}€\nEnvio: {envio_custo:.2f}€\nEstado de pagamento: {pagamento_status}\nTotal: {total + envio_custo:.2f}€\nObrigado pela sua compra!",
        "reserva_alterada": f"Olá {row['User']}, a sua reserva foi alterada.\n\n{itens_pedido}\n\nValor total dos artigos: {total:.2f}€\nEnvio: {envio_custo:.2f}€\nEstado de pagamento: {pagamento_status}\nTotal: {total + envio_custo:.2f}€\nObrigado pela sua compra!",
        "pagamento_recebido": f"Olá {row['User']}, recebemos o pagamento da sua reserva.\n\n{itens_pedido}\n\nValor total dos artigos: {total:.2f}€\nEnvio: {envio_custo:.2f}€\nEstado de pagamento: Pago\nTotal: {total + envio_custo:.2f}€\nObrigado pela sua compra!",
        "aviso_pagamento": f"Olá {row['User']}, os artigos reservados estão prestes a chegar.\n\n{itens_pedido}\n\nValor total dos artigos: {total:.2f}€\nEnvio: {envio_custo:.2f}€\nEstado de pagamento: {pagamento_status}\nTotal: {total + envio_custo:.2f}€\nObrigado pela sua compra!"
    }
    
    return mensagens.get(tipo_mensagem, None)

# Percorrer todas as folhas do Excel
for sheet_name, df in sheets.items():
    print(f"📄 Processando folha: {sheet_name}")

    if "User" not in df.columns or "Tipo de Mensagem" not in df.columns:
        print(f"❌ A coluna 'User' ou 'Tipo de Mensagem' não foi encontrada na folha {sheet_name}")
        continue

    for index, row in df.iterrows():
        username = row["User"].strip() if isinstance(row["User"], str) else ""
        enviado = row.get("Enviado ?", 0)
        tipo_mensagem = row.get("Tipo de Mensagem", "").strip().lower()

        # Debug para ver qual tipo de mensagem está sendo capturado
        print(f"📌 Usuário: {username} | Tipo de Mensagem: {tipo_mensagem} | Enviado: {enviado}")

        if tipo_mensagem not in ["nova_reserva", "reserva_alterada", "pagamento_recebido", "aviso_pagamento"]:
            print(f"❌ Tipo de mensagem inválido para {username}. Pulando...")
            continue

        if pd.isna(username) or username == "":
            continue

        if enviado == 1:
            print(f"✅ Mensagem já enviada para {username}, pulando...")
            continue  

        try:
            user_id = cl.user_id_from_username(username.replace("@", ""))
        except Exception as e:
            print(f"⚠️ Erro ao obter ID do usuário {username}: {e}. Pulando...")
            continue

        mensagem = criar_mensagem(row)
        if not mensagem:
            print(f"⚠️ Mensagem não gerada corretamente para {username}. Pulando...")
            continue

        print(f"📩 Enviando mensagem para {username}:\n{mensagem}")
        
        try:
            time.sleep(5)
            cl.direct_send(mensagem, [user_id])
            print(f"📩 Mensagem enviada para {username} (folha: {sheet_name})")
            df.at[index, "Enviado ?"] = 1
        except Exception as e:
            print(f"❌ Erro ao enviar mensagem para {username} (folha: {sheet_name}): {e}")

# Salvar o arquivo Excel atualizado
try:
    with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
        for sheet_name, df in sheets.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
except Exception as e:
    print(f"❌ Erro ao salvar o arquivo Excel: {e}")

# Logout após o envio
cl.logout()