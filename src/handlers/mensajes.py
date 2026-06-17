# import os
# import tempfile
# from estados import EstadoConversacion # manejo de estado por chat

# estado_chats = {} # {chat_id: EstadoConversacion}

# async def handle_message_received(evento: dict, wsp_client):
#     data = evento.get("data", {})
#     chat_id = data.get("chatId")
#     cuerpo = data.get("body", "").strip()
#     tipo = data.get("type")

#     if not cuerpo or tipo == "e2e_notification":
#         return

#     estado = estado_chats.get(chat_id)

#     # Comando inicial
#     if cuerpo.lower() == "!puntualidad":
#         estado_chats[chat_id] = EstadoConversacion(chat_id)
#         await wsp_client.enviar_mensaje(chat_id, 
#             "*Análisis de Puntualidad*\n\nPor favor envía el archivo *Anexo 5* (A5) en formato .xlsx")
#         return
    
#     if estado is None:
#         return #ignorar msj sin contexto
    
#     #recibir archivos adjuntos
#     if tipo in ("document", "image"):
#         archivo = data.get("mediaUrl") or data.get("body")

#         if not estado.tiene_a5:
#             estado.archivo_a5 = archivo
#             estado.tiene_a5 = True
#             await wsp_client.enviar_mensaje(chat_id, "Anexo 5 Recibido.\n\nAhora envía archivo de operación por favor (formato .xslx).")
#             return
        
#         if not estado.tiene_operacion:
#             estado.archivo_operacion = archivo
#             estado.tiene_operacion = True
#             await wsp_client.enviar_mensaje(chat_id, "Procesando datos...")

#             try:
#                 resultado = await procesar_puntualidad(estado)
#                 await wsp_client.enviar_mensaje(chat_id, resultado)

#             except Exception as e:
#                 await wsp_client.enviar_mensaje(chat_id, f"Error procesando los archivos: {e}")
            
#             finally:
#                 del estado_chats[chat_id]
