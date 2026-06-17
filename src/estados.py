class EstadoConversacion:
    def __init__(self, chat_id: str):
        self.chat_id = chat_id
        self.tiene_a5 = False
        self.tiene_operacion = False
        self.bytes_a5 = None
        self.bytes_operacion = None
        self.nombre_operacion = "operacion"