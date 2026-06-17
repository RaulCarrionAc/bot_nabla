import pandas as pd
import numpy as np
import io

def procesar_puntualidad(bytes_operacion: bytes, bytes_a5: bytes, nombre: str) -> tuple[bytes, str]:
    """Procesa los archivos y retorna (excel_bytes, resumen_texto)."""
    BD1 = pd.read_excel(io.BytesIO(bytes_operacion))  # Operacion
    BD2 = pd.read_excel(io.BytesIO(bytes_a5))         # Anexo5

    Bd_inicial_EX = BD1[["Fecha","Variante","Estado","Dirección","Tipo de Día","Período","01","Con Despacho Asociado"]]
    Bd_inicial_EX = Bd_inicial_EX.rename(columns={"Variante":"Servicio","Dirección":"Sentido"})
    Bd_inicial_A5 = BD2[["Servicio","Sentido","Anterior","Hora programada","Posterior","Tipo de Día"]]

    Bd_inicial_A5["ID"] = np.arange(1, len(Bd_inicial_A5)+1).astype(str)

    Bd_inicial_A5["Anterior"]        = pd.to_timedelta(Bd_inicial_A5["Anterior"].astype(str))
    Bd_inicial_A5["Hora programada"] = pd.to_timedelta(Bd_inicial_A5["Hora programada"].astype(str))
    Bd_inicial_A5["Posterior"]       = pd.to_timedelta(Bd_inicial_A5["Posterior"].astype(str))
    Bd_inicial_EX["01"]              = pd.to_timedelta(Bd_inicial_EX["01"].astype(str))
    Bd_inicial_EX["Fecha"]           = Bd_inicial_EX["Fecha"].astype(str)

    Bd_inicial_A5["Hora_anterior"] = Bd_inicial_A5["Hora programada"] - Bd_inicial_A5["Anterior"]
    Bd_inicial_A5["Hora_posterior"] = Bd_inicial_A5["Hora programada"] + Bd_inicial_A5["Posterior"]

    reemplazos = {
        "DLN":"DL","DOM":"DF","SAB":"DS",
        1:"Reg",0:"Ida",
        "R793_I":"R793","R796_I":"R796","R799_I":"R799",
        "R801_I":"R801","R800_R":"R800","R790V_R":"R790V",
    }
    Bd_inicial_EX["Tipo de Día"] = Bd_inicial_EX["Tipo de Día"].replace(reemplazos)
    Bd_inicial_EX["Servicio"]    = Bd_inicial_EX["Servicio"].replace(reemplazos)
    Bd_inicial_A5["Sentido"]     = Bd_inicial_A5["Sentido"].replace(reemplazos)

    for col in ["Servicio","Sentido","Tipo de Día"]:
        Bd_inicial_A5[col] = Bd_inicial_A5[col].astype(str)
        Bd_inicial_EX[col] = Bd_inicial_EX[col].astype(str)

    Bd_inicial_A5["key"] = Bd_inicial_A5["Servicio"]+Bd_inicial_A5["Sentido"]+Bd_inicial_A5["Tipo de Día"]+Bd_inicial_A5["ID"]
    Bd_not_match = Bd_inicial_EX.copy()
    Bd_inicial_EX = Bd_inicial_EX[Bd_inicial_EX["Estado"]=="Válida"]

    Bd_unida = Bd_inicial_EX.merge(
        Bd_inicial_A5[["Servicio","Sentido","Tipo de Día","Anterior","Hora programada",
                        "Posterior","Hora_anterior","Hora_posterior","ID","key"]],
        on=["Servicio","Sentido","Tipo de Día"], how="left"
    )
    Bd_unida["coincidencia"] = (
        (Bd_unida["01"] >= Bd_unida["Hora_anterior"]) &
        (Bd_unida["01"] <= Bd_unida["Hora_posterior"])
    )
    Bd_Filtrada = Bd_unida[Bd_unida["coincidencia"]].copy()

    # Franjas de puntualidad
    Bd_Filtrada["Anterior_0,25"]   = Bd_Filtrada["Hora programada"]-(Bd_Filtrada["Anterior"]/3)
    Bd_Filtrada["Anterior_0,25_2"] = Bd_Filtrada["Hora programada"]-(Bd_Filtrada["Anterior"]/4)
    Bd_Filtrada["Anterior_0,5"]    = Bd_Filtrada["Hora programada"]-(Bd_Filtrada["Anterior"]/4)
    Bd_Filtrada["Anterior_0,5_2"]  = Bd_Filtrada["Hora programada"]-(Bd_Filtrada["Anterior"]/6)
    Bd_Filtrada["Anterior_0,75"]   = Bd_Filtrada["Hora programada"]-(Bd_Filtrada["Anterior"]/6)
    Bd_Filtrada["Anterior_0,75_2"] = Bd_Filtrada["Hora programada"]-(Bd_Filtrada["Anterior"]/12)
    Bd_Filtrada["Anterior_1"]      = Bd_Filtrada["Hora programada"]-(Bd_Filtrada["Anterior"]/12)
    Bd_Filtrada["Posterior_1"]     = Bd_Filtrada["Hora programada"]+(Bd_Filtrada["Posterior"]/6)
    Bd_Filtrada["Posterior_0,75"]  = Bd_Filtrada["Hora programada"]+(Bd_Filtrada["Posterior"]/6)
    Bd_Filtrada["Posterior_0,75_2"]= Bd_Filtrada["Hora programada"]+(Bd_Filtrada["Posterior"]/3)
    Bd_Filtrada["Posterior_0,5"]   = Bd_Filtrada["Hora programada"]+(Bd_Filtrada["Posterior"]/3)
    Bd_Filtrada["Posterior_0,5_2"] = Bd_Filtrada["Hora programada"]+(Bd_Filtrada["Posterior"]/2)
    Bd_Filtrada["Posterior_0,25"]  = Bd_Filtrada["Hora programada"]+(Bd_Filtrada["Posterior"]/2)
    Bd_Filtrada["Posterior_0,25_2"]= Bd_Filtrada["Hora programada"]+(Bd_Filtrada["Posterior"]*(2/3))

    condiciones = [
        (Bd_Filtrada["01"]>=Bd_Filtrada["Anterior_1"]) & (Bd_Filtrada["01"]<=Bd_Filtrada["Posterior_1"]),
        ((Bd_Filtrada["01"]>=Bd_Filtrada["Anterior_0,75"]) & (Bd_Filtrada["01"]<Bd_Filtrada["Anterior_0,75_2"])) |
        ((Bd_Filtrada["01"]>Bd_Filtrada["Posterior_0,75"]) & (Bd_Filtrada["01"]<=Bd_Filtrada["Posterior_0,75_2"])),
        ((Bd_Filtrada["01"]>=Bd_Filtrada["Anterior_0,5"]) & (Bd_Filtrada["01"]<Bd_Filtrada["Anterior_0,5_2"])) |
        ((Bd_Filtrada["01"]>Bd_Filtrada["Posterior_0,5"]) & (Bd_Filtrada["01"]<=Bd_Filtrada["Posterior_0,5_2"])),
        ((Bd_Filtrada["01"]>=Bd_Filtrada["Anterior_0,25"]) & (Bd_Filtrada["01"]<Bd_Filtrada["Anterior_0,25_2"])) |
        ((Bd_Filtrada["01"]>Bd_Filtrada["Posterior_0,25"]) & (Bd_Filtrada["01"]<=Bd_Filtrada["Posterior_0,25_2"])),
    ]
    Bd_Filtrada["Indicador"] = np.select(condiciones, [1, 0.75, 0.5, 0.25], default=0.0)
    Bd_Filtrada["key2"] = Bd_Filtrada["Fecha"]+Bd_Filtrada["key"]
    Bd_Filtrada = Bd_Filtrada.sort_values(by=["key2","Indicador"], ascending=[True,False])
    Bd_Filtrada = Bd_Filtrada.drop_duplicates(subset=["key2"], keep="first")

    Bd_unida_not_match = Bd_not_match.merge(
        Bd_inicial_A5[["Servicio","Sentido","Tipo de Día","Anterior","Hora programada",
                        "Posterior","Hora_anterior","Hora_posterior","ID","key"]],
        on=["Servicio","Sentido","Tipo de Día"], how="left"
    )
    Bd_unida_not_match["key2"] = Bd_unida_not_match["Fecha"]+Bd_unida_not_match["key"]
    Bd_unida_not_match = Bd_unida_not_match.sort_values(by=["key2","01"])
    Bd_unida_not_match = Bd_unida_not_match.drop_duplicates(subset=["key2"], keep="first")
    Bd_unida_not_match["Indicador"] = 0

    Exp_FR = Bd_unida_not_match[~Bd_unida_not_match["key2"].isin(Bd_Filtrada["key2"])]
    Exp_FR = Exp_FR.dropna(subset=["key"])
    Bd_final = pd.concat([Bd_Filtrada, Exp_FR], ignore_index=True)
    Bd_final["Delta"] = abs(Bd_final["01"]-Bd_final["Hora programada"])

    # Franjas para Bd_final
    for col in ["Anterior_1","Posterior_1","Anterior_0,25","Anterior_0,75_2","Posterior_0,75","Posterior_0,25_2"]:
        if col not in Bd_final.columns:
            Bd_final[col] = pd.NaT

    condiciones2 = [
        (Bd_final["01"]>=Bd_final["Anterior_1"]) & (Bd_final["01"]<=Bd_final["Posterior_1"]),
        (Bd_final["01"]>=Bd_final["Anterior_0,25"]) & (Bd_final["01"]<=Bd_final["Anterior_0,75_2"]),
        (Bd_final["01"]>=Bd_final["Posterior_0,75"]) & (Bd_final["01"]<Bd_final["Posterior_0,25_2"]),
    ]
    Bd_final["Estatus"] = np.select(condiciones2, ["A tiempo","Adelantado","Atrasado"], default="Invalida/Fuera de Rango")

    Bd_final_v2 = Bd_final[["Fecha","Servicio","Sentido","Tipo de Día","Hora programada","Delta","Estatus","Indicador"]]

    # Resumen
    total = len(Bd_final_v2)
    indicador_prom = Bd_final_v2["Indicador"].mean() * 100
    dist = Bd_final_v2["Estatus"].value_counts().to_dict()

    resumen = (
        f"✅ *Resultado Puntualidad*\n\n"
        f"📋 Total registros: {total}\n"
        f"📊 Indicador promedio: *{indicador_prom:.1f}%*\n\n"
        f"🟢 A tiempo: {dist.get('A tiempo', 0)}\n"
        f"🔵 Adelantado: {dist.get('Adelantado', 0)}\n"
        f"🔴 Atrasado: {dist.get('Atrasado', 0)}\n"
        f"⚫ Fuera de rango: {dist.get('Invalida/Fuera de Rango', 0)}"
    )

    # Exportar a bytes
    output = io.BytesIO()
    Bd_final_v2.to_excel(output, index=False)
    excel_bytes = output.getvalue()

    return excel_bytes, resumen