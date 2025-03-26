def exibir_entidades(doc):
    """Exibe as entidades do DXF carregado."""
    if doc is None:
        print("Nenhum documento DXF carregado.")
        return
    
    modelspace = doc.modelspace()
    print("\nEntidades no arquivo DXF:")
    
    for entidade in modelspace:
        print(f"Tipo: {entidade.dxftype()} - Camada: {entidade.dxf.layer}")

    print("\nFim da exibição das entidades.")
