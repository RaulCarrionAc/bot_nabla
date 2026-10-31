def identificadores_remitente(data: dict, telefono_resuelto: str) -> set[str]:
    """Conserva teléfono y JID del remitente; nunca usa el chat como identidad."""
    sender = data.get("sender")
    sender_id = sender.get("id") if isinstance(sender, dict) else None
    original = data.get("author") or sender_id or data.get("from") or ""
    valores = [telefono_resuelto]
    # En grupos, `from` puede identificar el grupo, no a quien escribió.
    if not str(original).endswith("@g.us"):
        valores.append(str(original).split("@")[0].split(":")[0])
    return {limpio for valor in valores if (limpio := "".join(filter(str.isdigit, str(valor))))}
