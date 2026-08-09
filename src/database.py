import os
from typing import Optional
from datetime import datetime
from sqlmodel import SQLModel, Field, create_engine, Session, select
from sqlalchemy import Column, String, LargeBinary, Integer, Float

# Definir la URL de la base de datos (por defecto SQLite en un volumen compartido)
DATABASE_URL = os.getenv("DATABASE_URL", "sqlite:///data/bot_nabla.db")

# Asegurar que el directorio de la base de datos existe
if DATABASE_URL.startswith("sqlite:///"):
    db_file_path = DATABASE_URL.replace("sqlite:///", "")
    os.makedirs(os.path.dirname(os.path.abspath(db_file_path)), exist_ok=True)

# Crear el motor de la base de datos
engine = create_engine(DATABASE_URL, echo=False)

class ConversacionState(SQLModel, table=True):
    __tablename__ = "conversaciones"
    
    chat_id: str = Field(primary_key=True)
    estado: str
    tiene_a5: bool = Field(default=False)
    tiene_operacion: bool = Field(default=False)
    bytes_a5: Optional[bytes] = Field(sa_column=Column(LargeBinary, nullable=True))
    bytes_operacion: Optional[bytes] = Field(sa_column=Column(LargeBinary, nullable=True))
    nombre_operacion: str = Field(default="operacion")
    updated_at: datetime = Field(default_factory=datetime.utcnow)

class OperacionRecord(SQLModel, table=True):
    __tablename__ = "operaciones"
    
    id: Optional[int] = Field(default=None, primary_key=True)
    Fecha: str
    Variante: str
    Estado: str
    Dirección: str
    tipo_de_dia: str = Field(sa_column=Column("Tipo de Día", String))
    Período: str
    hora_01: str = Field(sa_column=Column("01", String))
    con_despacho_asociado: str = Field(sa_column=Column("Con Despacho Asociado", String))
    created_at: datetime = Field(default_factory=datetime.utcnow)

class Anexo5Record(SQLModel, table=True):
    __tablename__ = "anexo_5"
    
    id: Optional[int] = Field(default=None, primary_key=True)
    Servicio: str
    Sentido: str
    Anterior: str
    hora_programada: str = Field(sa_column=Column("Hora programada", String))
    Posterior: str
    tipo_de_dia: str = Field(sa_column=Column("Tipo de Día", String))
    created_at: datetime = Field(default_factory=datetime.utcnow)

class Expedicion(SQLModel, table=True):
    __tablename__ = "expediciones"
    
    id: Optional[int] = Field(default=None, primary_key=True)
    operador: str
    Fecha: str
    Inicio_Expedicion: str = Field(sa_column=Column("Inicio Expedicion", String))
    Fin_Expedicion: str = Field(sa_column=Column("Fin Expedicion", String))
    Folio_TS: Optional[str] = Field(default=None, sa_column=Column("Folio TS", String))
    ID_Exp: Optional[int] = Field(default=None, sa_column=Column("ID Exp", Integer))
    Bus: Optional[str] = Field(default=None)
    Chofer: Optional[str] = Field(default=None)
    Propietario: Optional[str] = Field(default=None)
    Variante: str
    Servicio: str
    Periodo: int
    Sentido: str
    Estado: Optional[str] = Field(default=None)
    Causa: Optional[str] = Field(default=None)
    Tipo_demanda: Optional[str] = Field(default=None, sa_column=Column("Tipo demanda", String))
    Frecuencia: Optional[float] = Field(default=None)
    Vel_Promedio: Optional[float] = Field(default=None, sa_column=Column("Vel.Promedio", Float))
    Vel_Maxima: Optional[float] = Field(default=None, sa_column=Column("Vel.Maxima", Float))
    
    # Checkpoints POIs 1 a 20 (Guardados como texto/hora)
    p1: Optional[str] = Field(default=None, sa_column=Column("1", String, nullable=True))
    p2: Optional[str] = Field(default=None, sa_column=Column("2", String, nullable=True))
    p3: Optional[str] = Field(default=None, sa_column=Column("3", String, nullable=True))
    p4: Optional[str] = Field(default=None, sa_column=Column("4", String, nullable=True))
    p5: Optional[str] = Field(default=None, sa_column=Column("5", String, nullable=True))
    p6: Optional[str] = Field(default=None, sa_column=Column("6", String, nullable=True))
    p7: Optional[str] = Field(default=None, sa_column=Column("7", String, nullable=True))
    p8: Optional[str] = Field(default=None, sa_column=Column("8", String, nullable=True))
    p9: Optional[str] = Field(default=None, sa_column=Column("9", String, nullable=True))
    p10: Optional[str] = Field(default=None, sa_column=Column("10", String, nullable=True))
    p11: Optional[str] = Field(default=None, sa_column=Column("11", String, nullable=True))
    p12: Optional[str] = Field(default=None, sa_column=Column("12", String, nullable=True))
    p13: Optional[str] = Field(default=None, sa_column=Column("13", String, nullable=True))
    p14: Optional[str] = Field(default=None, sa_column=Column("14", String, nullable=True))
    p15: Optional[str] = Field(default=None, sa_column=Column("15", String, nullable=True))
    p16: Optional[str] = Field(default=None, sa_column=Column("16", String, nullable=True))
    p17: Optional[str] = Field(default=None, sa_column=Column("17", String, nullable=True))
    p18: Optional[str] = Field(default=None, sa_column=Column("18", String, nullable=True))
    p19: Optional[str] = Field(default=None, sa_column=Column("19", String, nullable=True))
    p20: Optional[str] = Field(default=None, sa_column=Column("20", String, nullable=True))
    p21: Optional[str] = Field(default=None, sa_column=Column("21", String, nullable=True))
    p22: Optional[str] = Field(default=None, sa_column=Column("22", String, nullable=True))
    
    created_at: datetime = Field(default_factory=datetime.utcnow)

class Anexo1(SQLModel, table=True):
    __tablename__ = "anexo_1"
    
    id: Optional[int] = Field(default=None, primary_key=True)
    operador: str
    servicio: str
    sentido: str
    periodo: int
    horario: str
    tipo_dia: str = Field(sa_column=Column("Tipo de Día", String))
    tipo_demanda: Optional[str] = Field(default=None, sa_column=Column("Tipo Demanda", String))
    frecuencia_esperada: Optional[float] = Field(default=None, sa_column=Column("Frecuencia (buses/hr)", Float))
    created_at: datetime = Field(default_factory=datetime.utcnow)

class UsuarioPermitido(SQLModel, table=True):
    __tablename__ = "usuarios_permitidos"
    
    celular: str = Field(primary_key=True)
    created_at: datetime = Field(default_factory=datetime.utcnow)

class PuntoControlPO(SQLModel, table=True):
    __tablename__ = "puntos_control_po"
    
    id: Optional[int] = Field(default=None, primary_key=True)
    operador: str = Field(default="tasacop", index=True)
    servicio: str = Field(index=True)
    sentido: int = Field(description="0: Ida, 1: Regreso")
    correlativo: int = Field(description="1 a 22")
    distancia_origen: float = Field(description="Distancia en metros al origen")
    created_at: datetime = Field(default_factory=datetime.utcnow)


def init_db():
    """Crea las tablas en la base de datos si no existen."""
    SQLModel.metadata.create_all(engine)

def obtener_estado(chat_id: str) -> Optional[ConversacionState]:
    """Obtiene el estado de la conversación para un chat específico."""
    with Session(engine) as session:
        statement = select(ConversacionState).where(ConversacionState.chat_id == chat_id)
        results = session.exec(statement)
        return results.first()

def guardar_estado(
    chat_id: str, 
    estado: str, 
    tiene_a5: bool = False, 
    tiene_operacion: bool = False, 
    bytes_a5: bytes = None, 
    bytes_operacion: bytes = None, 
    nombre_operacion: str = 'operacion'
):
    """Guarda o actualiza el estado de la conversación para un chat."""
    with Session(engine) as session:
        state = session.get(ConversacionState, chat_id)
        if not state:
            state = ConversacionState(chat_id=chat_id, estado=estado)
        state.estado = estado
        state.tiene_a5 = tiene_a5
        state.tiene_operacion = tiene_operacion
        state.bytes_a5 = bytes_a5
        state.bytes_operacion = bytes_operacion
        state.nombre_operacion = nombre_operacion
        state.updated_at = datetime.utcnow()
        session.add(state)
        session.commit()

def eliminar_estado(chat_id: str):
    """Elimina el estado de la conversación para limpiar la interacción."""
    with Session(engine) as session:
        state = session.get(ConversacionState, chat_id)
        if state:
            session.delete(state)
            session.commit()

def es_usuario_permitido(celular: str) -> bool:
    """Verifica si un número de celular está en la tabla de permitidos."""
    with Session(engine) as session:
        user = session.get(UsuarioPermitido, celular)
        return user is not None

def agregar_usuario_permitido(celular: str) -> bool:
    """Agrega un número de celular a la tabla de permitidos."""
    with Session(engine) as session:
        user = session.get(UsuarioPermitido, celular)
        if not user:
            user = UsuarioPermitido(celular=celular)
            session.add(user)
            session.commit()
            return True
        return False

def quitar_usuario_permitido(celular: str) -> bool:
    """Remueve un número de celular de la tabla de permitidos."""
    with Session(engine) as session:
        user = session.get(UsuarioPermitido, celular)
        if user:
            session.delete(user)
            session.commit()
            return True
        return False

def listar_usuarios_permitidos() -> list[str]:
    """Retorna la lista de todos los celulares permitidos."""
    with Session(engine) as session:
        statement = select(UsuarioPermitido)
        results = session.exec(statement).all()
        return [user.celular for user in results]

