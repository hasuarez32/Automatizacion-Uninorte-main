def excel_exportar(data, nombre_archivo,numerodepoblacion, Preguntas,columnas_observaciones,general,oficina,proceso, perido,tipos_grafica, columnas_filtros_dinamicos=[], col_segmentacion="", poblaciones_por_segmento=None):
    #Hoja de Dijitación
    import xlsxwriter
    import pandas as pd
    import sys
    import math
    from xlsxwriter.utility import xl_rowcol_to_cell
    from xlsxwriter.utility import xl_cell_to_rowcol
    from datetime import datetime
    import datetime as dt
    from collections import Counter
    import numpy as np
    import textwrap

    workbook = xlsxwriter.Workbook(f'{nombre_archivo}.xlsx')

    # Variable para determinar si hay filtros dinámicos
    if poblaciones_por_segmento is None:
        poblaciones_por_segmento = {}
    tiene_filtros = len(columnas_filtros_dinamicos) > 0
    
    # Función helper para generar fórmulas que consideran filas visibles
    def countif_visible(columna, criterio):
        """Genera COUNTIF que considera solo filas visibles cuando hay slicers"""
        if tiene_filtros:
            return f'COUNTIFS(TB[{columna}],{criterio},TB[_VISIBLE],1)'
        else:
            return f'COUNTIF(TB[{columna}],{criterio})'

    def correl_visible(col_general, col_pregunta):
        """Correlacion Pearson solo para filas visibles (slicers activos)"""
        if tiene_filtros:
            g  = f'TB[{col_general}]'
            p  = f'TB[{col_pregunta}]'
            v  = 'TB[_VISIBLE]'
            n  = f'SUMPRODUCT({v})'
            sg = f'SUMPRODUCT(({v}=1)*{g})'
            sp = f'SUMPRODUCT(({v}=1)*{p})'
            sgp= f'SUMPRODUCT(({v}=1)*{g}*{p})'
            sg2= f'SUMPRODUCT(({v}=1)*{g}^2)'
            sp2= f'SUMPRODUCT(({v}=1)*{p}^2)'
            num= f'({sgp}-{sg}*{sp}/{n})'
            den= f'SQRT(({sg2}-{sg}^2/{n})*({sp2}-{sp}^2/{n}))'
            return f'=IFERROR(ABS({num}/{den}),0)'
        else:
            return f'=IFERROR(ABS(CORREL(TB[{col_general}], TB[{col_pregunta}])), 0)'
    Dijitacion = workbook.add_worksheet("Digitación")
    n_poblacion =numerodepoblacion
    n_estimado=data.shape[0]
    data = data.replace({np.nan: None, np.inf: None, -np.inf: None})
       # Escribir las cabeceras
    for col_num, header in enumerate(data.columns):
        Dijitacion.write(0, col_num, header)

    # Escribir los datos
    for row_num, row_data in enumerate(data.values):
        for col_num, cell_data in enumerate(row_data):
             # 1. Si es NaN o vacío → celda vacía
            if pd.isna(cell_data) or str(cell_data).strip() == '':
                Dijitacion.write_blank(row_num+1, col_num, None)

            # 2. Si es tipo fecha (datetime o fecha como string reconocida)
            elif isinstance(cell_data, (pd.Timestamp, dt.datetime, dt.date)):
                Dijitacion.write(row_num+1, col_num, cell_data.strftime('%d/%m/%Y'))

            elif isinstance(cell_data, str):
                try:
                    # Intentar convertir string a fecha
                    fecha = pd.to_datetime(cell_data, dayfirst=True, errors='raise')
                    Dijitacion.write_datetime(row_num+1, col_num, fecha.strftime('%d/%m/%Y'))
                except:
                    Dijitacion.write(row_num+1, col_num, cell_data)
            
            # 3. Otro tipo de dato
            else:
                Dijitacion.write(row_num+1, col_num, cell_data)

    n_rows, n_cols = data.shape

    # Si hay filtros, agregar columna auxiliar _VISIBLE
    if tiene_filtros:
        data['_VISIBLE'] = 1
        visible_col_idx = data.columns.get_loc('_VISIBLE')
        Dijitacion.write(0, visible_col_idx, '_VISIBLE')
        # Formula por fila: referencia directa a la celda (no TB[col])
        _q = chr(34)
        _fv_info_cols = []
        for _fi2, _fc2 in enumerate(columnas_filtros_dinamicos):
            if _fc2 in data.columns:
                _fa = xl_rowcol_to_cell(6 + _fi2 * 2, 1, row_abs=True, col_abs=True)
                _fci = data.columns.get_loc(_fc2)
                _fv_info_cols.append((_fa, _fci))
        def _make_vis_formula(_rn):
            if not _fv_info_cols:
                return '=1'
            _cs = [
                "OR('T+G'!" + _fa + '=' + _q + '(Todos)' + _q + ',' +
                xl_rowcol_to_cell(_rn + 1, _fci, col_abs=True) +
                "='T+G'!" + _fa + ')'
                for _fa, _fci in _fv_info_cols
            ]
            return '=IF(AND(' + ','.join(_cs) + '),1,0)'
        for row_num in range(len(data)):
            Dijitacion.write(row_num+1, visible_col_idx, _make_vis_formula(row_num))
        n_rows, n_cols = data.shape
        Dijitacion.add_table(0, 0, n_rows, n_cols - 1,
            {'columns': [{'header': col} for col in data.columns],
            'name': 'TB'})
        Dijitacion.set_column(visible_col_idx, visible_col_idx, None, None, {'hidden': True})
    else:
        Dijitacion.add_table(0, 0, n_rows, n_cols - 1,
            {'columns': [{'header': col} for col in data.columns],
            'name': 'TB'})
    for col_num, col_name in enumerate(data.columns):
        # Medir el ancho del header
        max_width = len(col_name)

        # Medir el ancho de cada celda en la columna
        for cell in data[col_name]:
            max_width = max(max_width, len(str(cell)))

        # Añadir un pequeño extra para que no quede muy justo
        Dijitacion.set_column(col_num, col_num, max_width + 2)

    # Crear formato de cabecera centrado
    header_format = workbook.add_format({
        'bold': True,        # Negrita
        'align': 'center',   # Centrar horizontalmente
        'valign': 'vcenter', # Centrar verticalmente
    
    })

    #formato de celda
    cell_format = workbook.add_format({
        'text_wrap': True,
        'align': 'center',
        'valign': 'vcenter',
        'bg_color': '#FFFFFF',  # Fondo blanco
        'border': 1             # Bordes
    })

    # Escribir los datos con formato de datos
    for row_num, row_data in enumerate(data.values):
        for col_num, cell_data in enumerate(row_data):
            # La columna _VISIBLE usa SUBTOTAL para detectar filas ocultas por el slicer;
            # no sobreescribir con el valor estático 1 del dataframe.
            if tiene_filtros and col_num == visible_col_idx:
                Dijitacion.write_formula(row_num+1, col_num, _make_vis_formula(row_num), cell_format)
                continue

            if pd.isna(cell_data) or str(cell_data).strip() == '':
                Dijitacion.write_blank(row_num+1, col_num, None,cell_format)

            # 2. Si es tipo fecha (datetime o fecha como string reconocida)
            elif isinstance(cell_data, (pd.Timestamp, dt.datetime, dt.date)):
                Dijitacion.write(row_num+1, col_num, cell_data.strftime('%d/%m/%Y'),cell_format)

            elif isinstance(cell_data, str):
                try:
                    # Intentar convertir string a fecha
                    fecha = pd.to_datetime(cell_data, dayfirst=True, errors='raise')
                    Dijitacion.write_datetime(row_num+1, col_num, fecha.strftime('%d/%m/%Y'),cell_format)
                except:
                    Dijitacion.write(row_num+1, col_num, cell_data,cell_format)

            # 3. Otro tipo de dato
            else:
                Dijitacion.write(row_num+1, col_num, cell_data,cell_format)

    # Escribir cabeceras con formato centrado
    for col_num, header in enumerate(data.columns):
        Dijitacion.write(0, col_num, header, header_format)

    #Función para calcular la poblacion estimada
    def calcular_poblacion_estimada(N):
        n=math.ceil((384.16)/(1+((384.16-1)/N)))
        return n
    
    #-------------------------------Labels-----------------------------------------------------------------------------------------------------------------------
    #Esquea de las preguntas 


    labels_graph = [ "Supera Notablemente las Expectativas", "4","3","2","Muy por Debajo de las Expectativas","No aplica"]
    respuestas = [5, 4,3,2,1,"No aplica"]

    #------------------------------------------------Generar grficas---------------------------------------------------------------------------------------------
    #Elaboracio ficha tecnica y grafica general
    #workbook = xlsxwriter.Workbook('combin.xlsx')
    TG = workbook.add_worksheet("T+G")
    Formato_ficha1 = workbook.add_format({
            'align': 'center',          # Alinear el texto al centro
            'left': 2,                  # Borde izquierdo grueso
            'right': 2,                 # Borde derecho grueso
            'top': 2,                   # Borde superior delgado
            'bottom': 2,                # Borde inferior delgado
            'bg_color': '#FABF8F',      # Color fondo gris  
            'border_color': 'black'     # Color del borde
        })
    Formato_ficha2 = workbook.add_format({
            'align': 'left',          # Alinear el texto al centro
            'left': 2,                  # Borde izquierdo grueso
            'right': 2,                 # Borde derecho grueso
            'top': 2,                   # Borde superior delgado
            'bottom': 2,                # Borde inferior delgado
            'bg_color': '#D3D3D3',      # Color fondo gris
            'border_color': 'black'     # Color del borde
        })
    Formato_ficha3 = workbook.add_format({
            'align': 'center',          # Alinear el texto al centro
            'left': 2,                  # Borde izquierdo grueso
            'right': 2,                 # Borde derecho grueso
            'top': 2,                   # Borde superior delgado
            'bottom': 2,                # Borde inferior delgado
            'bg_color': '#FFFFFF',      # Color fondo gris
            'border_color': 'black'     # Color del borde
        })
    Formato_ficha4 = workbook.add_format({
            'align': 'left',          # Alinear el texto al centro
            #'right': 2,                 # Borde derecho grueso
            #'top': 2,                   # Borde superior delgado
            #'bottom': 2,                # Borde inferior delgado
            'bg_color': '#D3D3D3',      # Color fondo gris
            'border_color': 'black'     # Color del borde
        })
    TG.merge_range(0, 1, 0, 22, 'EVALUACIÓN     DE     LA     CALIDAD     DE    LOS    SERVICIOS    EN    LA    ADMINISTRACIÓN    UNIVERSITARIA', Formato_ficha1)
    TG.merge_range(22, 1, 22, 22, 'S A T I S F A C C I Ó N     P O R     A T R I B U T O     E V A L U A D O', Formato_ficha1)
    TG.merge_range(1, 1, 1, 2, 'PERIODO', Formato_ficha2)
    TG.merge_range(2, 1, 2, 2, 'FECHA', Formato_ficha2)
    TG.merge_range(3, 1, 3, 2, 'OFICINA', Formato_ficha2)
    TG.merge_range(4, 1, 4, 2, 'PROCESO', Formato_ficha2)
    TG.merge_range(1, 3, 1, 10, f'{perido}', Formato_ficha3)
    TG.merge_range(2, 3, 2, 10, datetime.now().strftime("%d/%m/%Y"), Formato_ficha3)
    TG.merge_range(3, 3, 3, 10, f'{oficina}', Formato_ficha3)
    TG.merge_range(4, 3, 4, 10, f'{proceso}', Formato_ficha3)
    titulos_fichas = ["Población","Tamaño muestra estimada","Muestra alcanzada","Varianza estimada","Error estimado","Nivel de confabilidad"]

    for row in range(5, 19):  # filas 8 a 14
        for col in range(1, 11):  # columnas E a G
            TG.write(row, col, None, Formato_ficha4)
            if col == 1:
                TG.write(row, col,None, workbook.add_format({'align': 'left','left': 2, 'bg_color': '#D3D3D3','border_color': 'black'}))
            if row==18:
                TG.write(row, col,None, workbook.add_format({'align': 'left','bottom': 2, 'bg_color': '#D3D3D3','border_color': 'black'}))
                if col==1:
                    TG.write(row, col,None, workbook.add_format({'align': 'left','left': 2, 'bottom': 2,'bg_color': '#D3D3D3','border_color': 'black'}))
            if col==10:
                TG.write(row, col,None, workbook.add_format({'align': 'left','right': 2, 'bg_color': '#D3D3D3','border_color': 'black'}))
                if row==18:
                    TG.write(row, col,None, workbook.add_format({'align': 'left','right': 2, 'bottom': 2,'bg_color': '#D3D3D3','border_color': 'black'}))
    # ── Filtros desplegables (Data Validation nativa, sin VBA) ──────────
    if tiene_filtros:
        _fmt_lbl_f = workbook.add_format({
            'align': 'left', 'valign': 'vcenter', 'bold': True, 'font_size': 8,
            'bg_color': '#16365C', 'font_color': 'white',
            'left': 1, 'right': 1, 'top': 1, 'bottom': 0, 'border_color': 'black'
        })
        _fmt_lbl_pob = workbook.add_format({
            'align': 'left', 'valign': 'vcenter', 'bold': True, 'font_size': 8,
            'bg_color': '#375623', 'font_color': 'white',
            'left': 1, 'right': 1, 'top': 1, 'bottom': 0, 'border_color': 'black'
        })
        _fmt_val_f = workbook.add_format({
            'align': 'left', 'valign': 'vcenter', 'font_size': 9,
            'bg_color': '#FFFFFF', 'font_color': '#333333',
            'left': 1, 'right': 1, 'top': 0, 'bottom': 1, 'border_color': '#999999',
            'italic': True
        })
        for _fi, _fc in enumerate(columnas_filtros_dinamicos):
            _vals_dd = ['(Todos)'] + sorted(
                data[_fc].fillna('Sin especificar').astype(str).unique().tolist()
            )
            # Valores en columnas ocultas de T+G (col 30+, fila 1000+)
            _hcol = 30 + _fi
            _base_r = 1000
            for _rv, _vv in enumerate(_vals_dd):
                TG.write(_base_r + _rv, _hcol, _vv)
            TG.set_column(_hcol, _hcol, None, None, {'hidden': True})
            _fl_r = 5 + _fi * 2   # fila etiqueta (0-indexed)
            _fv_r = _fl_r + 1     # fila celda dropdown
            _es_pob = bool(col_segmentacion and _fc == col_segmentacion)
            _lbl_texto = 'Población' if _es_pob else _fc
            _fmt_lbl_uso = _fmt_lbl_pob if _es_pob else _fmt_lbl_f
            TG.merge_range(_fl_r, 1, _fl_r, 3, _lbl_texto, _fmt_lbl_uso)
            TG.merge_range(_fv_r, 1, _fv_r, 3, '(Todos)', _fmt_val_f)
            _v_s = xl_rowcol_to_cell(_base_r, _hcol, row_abs=True, col_abs=True)
            _v_e = xl_rowcol_to_cell(_base_r + len(_vals_dd) - 1, _hcol,
                                     row_abs=True, col_abs=True)
            TG.data_validation(_fv_r, 1, _fv_r, 3, {
                'validate': 'list',
                'source': f'={_v_s}:{_v_e}'
            })
        # ── Tabla lookup N por segmento de población ─────────────────────
        _seg_dropdown_cell = None
        _seg_lookup_range  = None
        if col_segmentacion and poblaciones_por_segmento and col_segmentacion in columnas_filtros_dinamicos:
            _idx_seg = columnas_filtros_dinamicos.index(col_segmentacion)
            _seg_dropdown_cell = xl_rowcol_to_cell(6 + _idx_seg * 2, 1, row_abs=True, col_abs=True)
            _seg_r0, _seg_ck, _seg_cv = 1200, 32, 33
            TG.write(_seg_r0, _seg_ck, '(Todos)')
            TG.write(_seg_r0, _seg_cv, n_poblacion)
            for _ii, (_sv, _sn) in enumerate(sorted(poblaciones_por_segmento.items())):
                TG.write(_seg_r0 + 1 + _ii, _seg_ck, str(_sv))
                TG.write(_seg_r0 + 1 + _ii, _seg_cv, int(_sn))
            TG.set_column(_seg_ck, _seg_cv, None, None, {'hidden': True})
            _seg_vs = xl_rowcol_to_cell(_seg_r0, _seg_ck, row_abs=True, col_abs=True)
            _seg_ve = xl_rowcol_to_cell(_seg_r0 + len(poblaciones_por_segmento), _seg_cv,
                                        row_abs=True, col_abs=True)
            _seg_lookup_range = f'{_seg_vs}:{_seg_ve}'
    else:
        _seg_dropdown_cell = None
        _seg_lookup_range  = None
    # ─────────────────────────────────────────────────────────────────────
    TG.merge_range(7,4, 7, 7,'FICHA TÉCNICA',workbook.add_format({'align': 'center',   'left': 2,'right': 2,'top': 2,'bottom': 1,'bg_color': '#D3D3D3','border_color': 'black', 'bold':True}))

    texto_mas_largo = max(titulos_fichas, key=len)
    ancho_aproximado = len(texto_mas_largo)*0.65
    TG.set_column(4, 5, ancho_aproximado)

    borde_personalizado = workbook.add_format({
            'align': 'center',          # Alinear el texto al centro
            'left': 2,                  # Borde izquierdo grueso
            'right': 2,                 # Borde derecho grueso
            'top': 2,                   # Borde superior delgado
            'bottom': 2,                # Borde inferior delgado
            'bg_color': '#D3D3D3',      # Color fondo gris
            'border_color': 'black'     # Color del borde
        })
    formato_combinado = workbook.add_format({
            'border': 2,                # Borde delgado alrededor de las celdas combinadas
            'align': 'center',          # Alinear el texto al centro
            'valign': 'vcenter',        # Alinear el texto verticalmente al centro
            'bg_color': '#16365C',      # Color de fondo azul
            'bold': True,               # Texto en negrita
            'border_color': 'black',    # Color del borde
            'font_color': 'white',       # Color de fuente blanco
            'text_wrap': True
        })
    formato_combinado2 = workbook.add_format({
            'border': 2,                # Borde delgado alrededor de las celdas combinadas
            'align': 'center',          # Alinear el texto al centro
            'valign': 'vcenter',        # Alinear el texto verticalmente al centro
            'bold': True,               # Texto en negrita
            'border_color': 'black',    # Color del borde
            'font_color': 'black'       # Color de fuente negro
        })
    formato_combinado3 = workbook.add_format({
            'border': 2,                # Borde delgado alrededor de las celdas combinadas
            'align': 'center',          # Alinear el texto al centro
            'valign': 'vcenter',        # Alinear el texto verticalmente al centro
            #'bold': True,               # Texto en negrita
            'border_color': 'black',    # Color del borde
            'bg_color': '#C5D9F1',      # Color fondo azul claro
            'font_color': 'black',       # Color de fuente blanco
            'text_wrap': True,
            'font_size': 9
        })
    formato_borde_personalizado = workbook.add_format({
            'num_format': '0%',  # Formato de porcentaje con dos decimales
            'left': 2,                  # Borde izquierdo grueso
            'right': 2,                 # Borde derecho grueso
            'top': 2,                   # Borde superior delgado
            'bottom': 2,                # Borde inferior delgado
            'align': 'center',          # Alinear el texto al centro
            'border_color': 'black'     # Color del borde
        })
    formato_borde_personalizado1 = workbook.add_format({
            'num_format': '0',  # Formato de porcentaje con dos decimales
            'left': 2,                  # Borde izquierdo grueso
            'right': 2,                 # Borde derecho grueso
            'top': 2,                   # Borde superior delgado
            'bottom': 2,                # Borde inferior delgado
            'align': 'center',          # Alinear el texto al centro
            'border_color': 'black'     # Color del borde
        })
    #Generar la grfica general
    TG.merge_range(3, 11, 3, 16, 'ATRIBUTO', borde_personalizado)
    TG.merge_range(3, 17, 3, 22, 'SATISFACIÓN', borde_personalizado)
    TG.merge_range(4, 11, 7, 16, f'{general}', formato_combinado)
    TG.merge_range(8, 11, 18, 22, '', formato_combinado2)
    for col in range(17, 23):
        TG.merge_range(4, col, 5, col, labels_graph[col-17], formato_combinado3)
        # Aplicar bordes a la fila siguiente en el mismo rango de columnas
    denominador = '$G$11' if tiene_filtros else str(n_estimado)
    for col in range(17, 23):
        if col==22:
            TG.write_formula(6, col, f'={countif_visible(general,"No Aplica")}/{denominador}', formato_borde_personalizado)
        else:
            TG.write_formula(6, col,f'={countif_visible(general,respuestas[col-17])}/({denominador}-{countif_visible(general,"No Aplica")})', formato_borde_personalizado)

    for col in range(17, 23):
        if col==22:
            TG.write_formula(7, col, f'={countif_visible(general,"No Aplica")}', formato_borde_personalizado1)
        else:
            TG.write_formula(7, col,f'={countif_visible(general,respuestas[col-17])}', formato_borde_personalizado1)

    #Grafica geeneral
    # Crea un gráfico de tipo columna en el bloque de la izquierda
        chart = workbook.add_chart({'type': 'column'})

        # Añade la serie
        chart.add_series({
            'categories': ['T+G', 4, 17, 5, 21],
            'values':     ['T+G', 6, 17, 6, 21],
            'data_labels': {
                'value': True,          # Muestra el valor
                'position': 'outside_end'  # etiquetas encima de las barras
            }
        })

        chart.set_legend({'none': True})
        chart.set_plotarea({'border': {'none': True}})

        chart.set_plotarea({
            'border': {'none': True},
            'fill':   {'none': True}
        })

        chart.set_chartarea({
            'border': {'none': True},
            'fill':   {'none': True}
        })

        # Además, desactiva todas las cuadrículas posibles:
        chart.set_x_axis({'major_gridlines': {'visible': False}})
        chart.set_y_axis({
            'major_gridlines': {'visible': False},
            'visible': False
        })

        # Inserta el gráfico
        TG.insert_chart(8, 11, chart, {
            'x_offset': 2.2,
            'y_offset': 2,
            'x_scale': 1.9,
            'y_scale': 0.612
        })
    #------------------------------------------------Aplicar formato---------------------------------------------------------------------------------------------
    
    # Definir la función para aplicar el formato
    def aplicar_formato(worksheet, start_row,deteccion):
        # Definir formatos
        formato_combinado = workbook.add_format({
            'border': 2,                # Borde delgado alrededor de las celdas combinadas
            'align': 'center',          # Alinear el texto al centro
            'valign': 'vcenter',        # Alinear el texto verticalmente al centro
            'bg_color': '#16365C',      # Color de fondo azul
            'bold': True,               # Texto en negrita
            'border_color': 'black',    # Color del borde
            'font_color': 'white',       # Color de fuente blanco
            'text_wrap': True
        })
        
        formato_combinado2 = workbook.add_format({
            'border': 2,                # Borde delgado alrededor de las celdas combinadas
            'align': 'center',          # Alinear el texto al centro
            'valign': 'vcenter',        # Alinear el texto verticalmente al centro
            'bold': True,               # Texto en negrita
            'border_color': 'black',    # Color del borde
            'font_color': 'black'       # Color de fuente negro
        })
        
        formato_combinado3 = workbook.add_format({
            'border': 2,                # Borde delgado alrededor de las celdas combinadas
            'align': 'center',          # Alinear el texto al centro
            'valign': 'vcenter',        # Alinear el texto verticalmente al centro
            #'bold': True,               # Texto en negrita
            'border_color': 'black',    # Color del borde
            'bg_color': '#C5D9F1',      # Color fondo azul claro
            'font_color': 'black',       # Color de fuente blanco
            'text_wrap': True,
            'font_size': 9
        })
        
        borde_personalizado = workbook.add_format({
            'align': 'center',          # Alinear el texto al centro
            'left': 2,                  # Borde izquierdo grueso
            'right': 2,                 # Borde derecho grueso
            'top': 2,                   # Borde superior delgado
            'bottom': 2,                # Borde inferior delgado
            'bg_color': '#D3D3D3',      # Color fondo gris
            'border_color': 'black'     # Color del borde
        })
        
        formato_borde_personalizado = workbook.add_format({
            'num_format': '0%',  # Formato de porcentaje con dos decimales
            'left': 2,                  # Borde izquierdo grueso
            'right': 2,                 # Borde derecho grueso
            'top': 2,                   # Borde superior delgado
            'bottom': 2,                # Borde inferior delgado
            'align': 'center',          # Alinear el texto al centro
            'border_color': 'black'     # Color del borde
        })
        
        #indice lista calculo
        k=int((start_row-24)/14)
        # Combinar celdas y aplicar formatos
        worksheet.merge_range(start_row, 1, start_row, 4, 'ATRIBUTO', borde_personalizado)
        worksheet.merge_range(start_row, 5, start_row, 10, 'SATISFACIÓN', borde_personalizado)
        worksheet.merge_range(start_row + 1, 1, start_row + 3, 4, Preguntas[2*k], formato_combinado)
        worksheet.merge_range(start_row + 4, 1, start_row + 12, 10, '', formato_combinado2)

        # Combinar celdas en columnas F a K (columnas 5 a 10)
        for col in range(5, 11):
            worksheet.merge_range(start_row + 1, col, start_row + 2, col, labels_graph[col-5], formato_combinado3)
            texto_mas_largo = max(labels_graph, key=len)
            ancho_aproximado = len(texto_mas_largo)*0.6
            worksheet.set_column(5,5, ancho_aproximado)
            worksheet.set_column(9,9, ancho_aproximado)

        # Aplicar bordes a la fila siguiente en el mismo rango de columnas
        denominador_attr = '$G$11' if tiene_filtros else str(n_estimado)
        for col in range(5, 11):
            if col==10:
                worksheet.write_formula(start_row + 3, col, f'={countif_visible(Preguntas[2*k], chr(34)+"No Aplica"+chr(34))}/{denominador_attr}', formato_borde_personalizado)
            else:
                worksheet.write_formula(start_row + 3, col, f'={countif_visible(Preguntas[2*k], respuestas[col-5])}/({denominador_attr}-{countif_visible(Preguntas[2*k], chr(34)+"No Aplica"+chr(34))})', formato_borde_personalizado)
        
        inicio_nsp=xl_rowcol_to_cell(start_row+3, 5)
        fin_nsp=xl_rowcol_to_cell(start_row+3, 10)
        range_nsp=[f'{inicio_nsp}:{fin_nsp}']
    # Crea un gráfico de tipo columna en el bloque de la izquierda
        chart = workbook.add_chart({'type': 'column'})

        # Añade la serie
        chart.add_series({
            'categories': ['T+G', start_row+1, 5, start_row+2, 9],
            'values':     ['T+G', start_row+3, 5, start_row+3, 9],
            'data_labels': {
                'value': True,          # Muestra el valor
                'position': 'outside_end'  # etiquetas encima de las barras
            }
        })

        chart.set_legend({'none': True})
        chart.set_plotarea({'border': {'none': True}})

        chart.set_plotarea({
            'border': {'none': True},
            'fill':   {'none': True}
        })

        chart.set_chartarea({
            'border': {'none': True},
            'fill':   {'none': True}
        })

        # Además, desactiva todas las cuadrículas posibles:
        chart.set_x_axis({'major_gridlines': {'visible': False}})
        chart.set_y_axis({
            'major_gridlines': {'visible': False},
            'visible': False
        })

        # Inserta el gráfico
        TG.insert_chart(start_row+4, 1, chart, {
            'x_offset': 2,
            'y_offset': 2,
            'x_scale': 1.8,
            'y_scale': 0.612
        })
        if start_row-24 <deteccion:
            # Combinar celdas y aplicar formatos derecha
            worksheet.merge_range(start_row, 12, start_row, 16, 'ATRIBUTO', borde_personalizado)
            worksheet.merge_range(start_row, 17, start_row, 22, 'SATISFACIÓN', borde_personalizado)
            worksheet.merge_range(start_row + 1, 12, start_row + 3, 16, Preguntas[2*k+1], formato_combinado)
            worksheet.merge_range(start_row + 4, 12, start_row + 12, 22, '', formato_combinado2)

            # Aplicar bordes a la fila siguiente en el mismo rango de columnas
            denominador_attr_r = '$G$11' if tiene_filtros else str(n_estimado)
            for col in range(17, 23):
                if col==22:
                    worksheet.write_formula(start_row + 3, col, f'={countif_visible(Preguntas[2*k+1], chr(34)+"No Aplica"+chr(34))}/{denominador_attr_r}', formato_borde_personalizado)
                else:
                    worksheet.write_formula(start_row + 3, col, f'={countif_visible(Preguntas[2*k+1], respuestas[col-17])}/({denominador_attr_r}-{countif_visible(Preguntas[2*k+1], chr(34)+"No Aplica"+chr(34))})', formato_borde_personalizado)
            
            for col in range(17, 23):
                worksheet.merge_range(start_row + 1, col, start_row + 2, col, labels_graph[col-17], formato_combinado3)
                texto_mas_largo = max(labels_graph, key=len)
                ancho_aproximado = len(texto_mas_largo)*0.6
                worksheet.set_column(17,17, ancho_aproximado)
                worksheet.set_column(21,21, ancho_aproximado)
            

            inicio_nsp=xl_rowcol_to_cell(start_row+3, 17)
            fin_nsp=xl_rowcol_to_cell(start_row+3, 22)
            range_nsp.append(f'{inicio_nsp}:{fin_nsp}')
            # Crea un gráfico de tipo columna en el bloque de la derecha
            chart = workbook.add_chart({'type': 'column'})

            # Añade la serie
            chart.add_series({
                'categories': ['T+G', start_row+1, 17, start_row+2, 21],
                'values':     ['T+G', start_row+3, 17, start_row+3, 21],
                'data_labels': {
                    'value': True,          # Muestra el valor
                    'position': 'outside_end'  # etiquetas encima de las barras
                }
            })

            chart.set_legend({'none': True})
            chart.set_plotarea({'border': {'none': True}})

            chart.set_plotarea({
                'border': {'none': True},
                'fill':   {'none': True}
            })

            chart.set_chartarea({
                'border': {'none': True},
                'fill':   {'none': True}
            })

            # Además, desactiva todas las cuadrículas posibles:
            chart.set_x_axis({'major_gridlines': {'visible': False}})
            chart.set_y_axis({
                'major_gridlines': {'visible': False},
                'visible': False
            })
            # Inserta el gráfico
            TG.insert_chart(start_row+4, 12, chart, {
                'x_offset': 2,
                'y_offset': 2,
                'x_scale': 1.8,
                'y_scale': 0.612
            })
        return range_nsp


    def paroimpar(n):
        if n % 2 == 0:
            return "Par"
        else:
            return "Impar"
        
    # Aplicar el formato cada 14 filas
    num_columas=len(Preguntas)+1   
    pi=paroimpar(num_columas)
    if pi == "Impar":
        num_repeticiones = int((num_columas-1)/2)
    else:
        num_repeticiones = int((num_columas)/2)
    Lista_rango_nsp=[]
    for i in range(num_repeticiones):
        start_row = i * 14  # Cada bloque comienza cada 14 filas
        if pi=="Par":
            deteccion=(num_repeticiones-1)*14
        else:
            deteccion=start_row+25
        Lista_rango_nsp.extend(aplicar_formato(TG, start_row+24, deteccion))


#-------------------------------------------------insertar valores ficha varianza----------------------------------------------------------------------------------
    # Insertar la fórmula ficha técnica (varianza estimada)
    formulas = []
    for v in range(len(Lista_rango_nsp)):
        inicio_ficha, fin_ficha = Lista_rango_nsp[v].split(':')
        ficha_ini, colficha_ini = xl_cell_to_rowcol(inicio_ficha) 
        ficha_fin, colficha_fin = xl_cell_to_rowcol(fin_ficha)    
        #Pesos correspondintes
        Suma1 = [xl_rowcol_to_cell(ficha_ini, col) for col in range(colficha_ini,colficha_fin-3)]
        Suma2 = [xl_rowcol_to_cell(ficha_ini, col) for col in range(colficha_ini+2,colficha_fin)]
        formula_v = f"(SUM({Suma1[0]}:{Suma1[-1]})*SUM({Suma2[0]}:{Suma2[-1]}))"
        formulas.append(formula_v)
    formula_varianza = f"=AVERAGE({','.join(formulas)})"
    for row in range(8, 14):
        if row==13:
            TG.merge_range(row, 4, row, 5, titulos_fichas[row-8], workbook.add_format({'align': 'center','left':2,'bottom':2,'bg_color': '#FFFFFF','border_color': 'black'}))
            TG.merge_range(row, 6, row, 7, '=IFERROR((1 - (((G9 - G11) / G9)^0.5) * ((G12 / (G11 - 1))^0.5)) * 100, 0)', workbook.add_format({'align': 'center','left':1,'right':2,'bottom':2,'bg_color': '#FFFFFF','border_color': 'black','num_format': '0.0'}))
        elif row==12:
            TG.merge_range(row, 4, row, 5, titulos_fichas[row-8], workbook.add_format({'align': 'center','left':2,'bottom':1,'bg_color': '#FFFFFF','border_color': 'black'}))
            TG.merge_range(row, 6, row, 7, '=1-G14%', workbook.add_format({'align': 'center','right':2,'left':1,'bottom':1,'bg_color': '#FFFFFF','border_color': 'black', 'num_format': '0.0%'}))
        elif row==11:
            TG.merge_range(row, 4, row, 5, titulos_fichas[row-8], workbook.add_format({'align': 'center','left':2,'bottom':1,'bg_color': '#FFFFFF','border_color': 'black'}))
            TG.merge_range(row, 6, row, 7, formula_varianza, workbook.add_format({'align': 'center','right':2,'left':1,'bottom':1,'bg_color': '#FFFFFF','border_color': 'black', 'num_format': '0.0%'}))
        elif row==10:
            TG.merge_range(row, 4, row, 5, titulos_fichas[row-8], workbook.add_format({'align': 'center','left':2,'bottom':1,'bg_color': '#FFFFFF','border_color': 'black'}))
            TG.merge_range(row, 6, row, 7, '=SUMPRODUCT(TB[_VISIBLE])' if tiene_filtros else data.shape[0], workbook.add_format({'align': 'center','right':2,'left':1,'bottom':1,'bg_color': '#FFFFFF','border_color': 'black', 'num_format': '0'}))
        elif row==9:
            TG.merge_range(row, 4, row, 5, titulos_fichas[row-8], workbook.add_format({'align': 'center','left':2,'bottom':1,'bg_color': '#FFFFFF','border_color': 'black'}))
            if _seg_lookup_range:
                _g_pob = xl_rowcol_to_cell(8, 6, row_abs=True, col_abs=True)
                _muestra_val = f'=ROUNDUP(384.16/(1+383.16/{_g_pob}),0)'
            else:
                _muestra_val = calcular_poblacion_estimada(n_poblacion)
            TG.merge_range(row, 6, row, 7, _muestra_val,
                workbook.add_format({'align': 'center','right':2,'left':1,'bottom':1,'bg_color': '#FFFFFF','border_color': 'black', 'num_format': '0'}))
        elif row==8:
            TG.merge_range(row, 4, row, 5, titulos_fichas[row-8], workbook.add_format({'align': 'center','left':2,'bottom':1,'bg_color': '#FFFFFF','border_color': 'black'}))
            if _seg_lookup_range and _seg_dropdown_cell:
                _g8_f = f'=IFERROR(VLOOKUP({_seg_dropdown_cell},{_seg_lookup_range},2,FALSE),{n_poblacion})'
                TG.merge_range(row, 6, row, 7, _g8_f, workbook.add_format({'align': 'center','right':2,'left':1,'bottom':1,'bg_color': '#FFFFFF','border_color': 'black', 'num_format': '0'}))
            else:
                TG.merge_range(row, 6, row, 7, n_poblacion, workbook.add_format({'align': 'center','right':2,'left':1,'bottom':1,'bg_color': '#FFFFFF','border_color': 'black', 'num_format': '0'}))
        else:
            TG.merge_range(row, 4, row, 5, titulos_fichas[row-8], workbook.add_format({'align': 'center','left':2,'bottom':1,'bg_color': '#FFFFFF','border_color': 'black'}))
            TG.merge_range(row, 6, row, 7, None, workbook.add_format({'align': 'center','right':2,'left':1,'bottom':1,'bg_color': '#FFFFFF','border_color': 'black'}))

    TG.merge_range(15,3,15,8,'CÁLCULO NIVEL DE SATISFACCIÓN GENERAL',workbook.add_format({'align': 'center',   'left': 2,'right': 2,'top': 2,'bottom': 2,'bg_color': '#D3D3D3','border_color': 'black', 'bold':True}))
    TG.merge_range(16,3,16,7,'Nivel de Satisfacción General (Escala 0:100)',workbook.add_format({'align': 'center', 'bg_color': '#C4D79B', 'left':2,'right':1,'top':2,'bottom':1,'border_color':'black','num_format':'0.00'}))
    TG.merge_range(17,3,17,7,'Nivel de Satisfacción General (Escala 0:5)',workbook.add_format({'align': 'center', 'bg_color': '#C4D79B', 'left':2,'right':1,'top':1,'bottom':2,'border_color':'black','num_format':'0.00'}))
    TG.write_formula(16, 8, '=(R7*100)+(S7*80)+(T7*60)+(U7*40)+(V7*20)', workbook.add_format({'align': 'center','right':2,'left':1,'top':2,'bottom':1,'bg_color': '#C4D79B','border_color':'black','num_format':'0.00'}))
    TG.write_formula(17, 8, '=I17/20', workbook.add_format({'align': 'center','right':2,'left':1,'top':1,'bottom':2,'bg_color': '#C4D79B','border_color':'black','num_format':'0.00'}))

#-------------------------------------------------elaboracion matriz satisfacioón----------------------------------------------------------------------------------
    #Elaboracion matiz de satifaccion e importanacia
    start_sati_impor=(num_repeticiones)*14+25
    id_row_peso=start_sati_impor
    id__row_nip=start_sati_impor
    id_row_nsp1=start_sati_impor
    id_row_nsp2=start_sati_impor
    id_row_pnsp=start_sati_impor
    id_row_codigo=start_sati_impor
    id_row_grafica=start_sati_impor+1
    id_row_isc=start_sati_impor
    TG.merge_range(start_sati_impor, 1, start_sati_impor, 22, 'S A T I S F A C C I Ó N     E    I M P O R T A N C I A     P O N D E R A D A ', Formato_ficha1)
    for f in range(1,22+1):
        for h in range(start_sati_impor+1, start_sati_impor+int(len(Preguntas))+19):
            if f==1:
                TG.write(h, f, None, workbook.add_format({'align': 'center', 'bg_color': '#FFFFFF', 'left':2,'border_color':'black'}))
                if h==start_sati_impor+int(len(Preguntas))+18:
                    TG.write(h, f, None, workbook.add_format({'align': 'center', 'bg_color': '#FFFFFF', 'left':2,'bottom':2,'border_color':'black'}))
            elif f==22:
                TG.write(h, f, None, workbook.add_format({'align': 'center', 'bg_color': '#FFFFFF', 'right':2,'border_color':'black'}))
                if h==start_sati_impor+int(len(Preguntas))+18:
                    TG.write(h, f, None, workbook.add_format({'align': 'center', 'bg_color': '#FFFFFF', 'right':2,'bottom':2,'border_color':'black'}))
            elif h==start_sati_impor+int(len(Preguntas))+18:
                TG.write(h, f, None, workbook.add_format({'align': 'center', 'bg_color': '#FFFFFF', 'bottom':2,'border_color':'black'}))
            else:
                TG.write(h, f, None, workbook.add_format({'align': 'center', 'bg_color': '#FFFFFF'}))
    TG.merge_range(start_sati_impor+int(len(Preguntas))+22, 1, start_sati_impor+int(len(Preguntas))+22, 22, 'ANALISIS GENERAL DE SATISFACCIÓN DE CLIENTES', 
                workbook.add_format({'align': 'center', 'bg_color': '#FFFFFF', 'left':2,'right':2,'top':2,'bottom':2,'border_color':'black'}))
    #-----------------------------------------------------------------------------funcion grafica------------------------------------------
    from collections import Counter
    k_index=1
    control_renglon=0
    c_index=0
    d_index=1
    hidden_sheet = workbook.add_worksheet('datos_para_grafica')
    for key,value in tipos_grafica.items():
        columna_pregunta = f"{key}"
        if columna_pregunta in data.columns:
            control_renglon=control_renglon+1
            # Contar las respuestas válidas (sin nulos)
            conteo_torta = Counter(data[columna_pregunta].dropna())
            total_torta = sum(conteo_torta.values())
            # Escribir encabezados en columnas más a la derecha (por ejemplo columna G = col 6)
            col_offset = 2+d_index  # columna G
            hidden_sheet.write(0, col_offset, 'Respuesta')
            hidden_sheet.write(0, col_offset + 1, 'Proporción')

            # Escribir datos en columnas G y H
            for j, (respuesta, frecuencia) in enumerate(conteo_torta.items(), start=1):
                hidden_sheet.write(j, col_offset, respuesta)
                hidden_sheet.write_number(j, col_offset + 1, frecuencia / total_torta)

            # Crear gráfico de torta
            grafico_analisis = workbook.add_chart({'type': f'{value}'})

            if value == 'column':
                grafico_analisis.add_series({
                    'name': f'{columna_pregunta}',
                    'categories': [hidden_sheet.name, 1, col_offset, j, col_offset],
                    'values':     [hidden_sheet.name, 1, col_offset + 1, j, col_offset + 1],
                    'fill':       {'color': "#5E6BB5"},  # Naranja
                    'data_labels': {'value': True,'num_format': '0.0%'},
                })
                grafico_analisis.set_style(2)
                grafico_analisis.set_legend({'none': True})
            else:
                grafico_analisis.add_series({
                    'name': f'{columna_pregunta}',
                    'categories': [hidden_sheet.name, 1, col_offset, j, col_offset],
                    'values':     [hidden_sheet.name, 1, col_offset + 1, j, col_offset + 1],
                    'data_labels': {'value': True,'num_format': '0.0%'},
                    'data_labels': {
                    'percentage': True,
                    'leader_lines': True
                },
                })
                grafico_analisis.set_style(10)
            grafico_analisis.set_x_axis({ 'major_gridlines': {'visible': False}})
            grafico_analisis.set_y_axis({'visible': False, 'major_gridlines': {'visible': False},'num_format': '0%'})
            grafico_analisis.set_plotarea({'border': {'none': True}})
            grafico_analisis.set_chartarea({'border': {'none': True}})
            if len(set(data[columna_pregunta].tolist()))>15 and value == 'column':
                grafico_analisis.set_size({'width': 1500, 'height': 400})
                validar_tamaño= True
                k_index=1
                if control_renglon % 3 == 0:
                        c_index=c_index+17
                if control_renglon == 2:
                    c_index=c_index+17
                c_index=c_index+17
            else:
                grafico_analisis.set_size({'width': 500, 'height': 300})
                validar_tamaño= False
            grafico_analisis.set_chartarea({'fill': {'none': True}, 'border': {'none': True}})
            grafico_analisis.set_plotarea({'border': {'none': True}, 'fill': {'none': True}})
            grafico_analisis.set_title({
                'name': f'{columna_pregunta}',
                'name_font': {
                    'bold': False,
                    'color': '#333333',
                    'size': 11
                    }
                })
            
            TG.insert_chart(start_sati_impor+int(len(Preguntas))+24+c_index, k_index, grafico_analisis, {'positioning': 1, 'object_position': 1})   
            k_index=k_index+8
            d_index=d_index+8
            if validar_tamaño== False:
                if control_renglon % 3 == 0:
                    if control_renglon == len(tipos_grafica):
                        k_index=k_index
                        c_index=c_index
                    else:
                        k_index = 1
                        c_index=c_index+17
                # else:
                #     if control_renglon==4:
                #         c_index=c_index+17
            else:
                if control_renglon % 3 == 0:
                      c_index=c_index+17
                      if len(set(data[columna_pregunta].tolist()))>15 and value == 'column':
                                k_index = 1
                                control_renglon = control_renglon -1
                                c_index = c_index + 2                       
                else:
                     c_index=c_index+17
                     if len(set(data[columna_pregunta].tolist()))>15 and value == 'column':
                                k_index = 1
                                control_renglon = control_renglon -1
                                c_index = c_index + 2
    hidden_sheet.hide()
    TG.merge_range(start_sati_impor+int(len(Preguntas))+23, 1, start_sati_impor+int(len(Preguntas))+40+c_index, 22, None, 
                workbook.add_format({'align': 'center', 'bg_color': '#FFFFFF', 'left':2,'right':2,'top':2,'bottom':2,'border_color':'black'}))
    
    #----------------------------------------------------------------------------------------formatos--------------------------------------------------------
    Formato_satis_import1  = workbook.add_format({
            'align': 'center',          # Alinear el texto al centro
            'left': 2,                  # Borde izquierdo grueso
            'right': 2,                 # Borde derecho grueso
            'top': 2,                   # Borde superior delgado
            'bottom': 2,                # Borde inferior delgado
            'bg_color': '#D3D3D3',      # Color fondo gris
            'border_color': 'black',     # Color del borde,
            'valign': 'vcenter',        # Alinear el texto verticalmente al centro
        })
    Formato_satis_import3 = workbook.add_format({
            'num_format': '0.00',  # Formato de porcentaje con dos decimales
            'align': 'Center',          # Alinear el texto al centro
            'left': 2,                  # Borde izquierdo grueso
            'right': 1,                 # Borde derecho grueso
            'top': 2,                   # Borde superior delgado
            'bottom': 2,                # Borde inferior delgado
            'bg_color': '#FFFFFF',      # Color fondo gris
            'border_color': 'black',     # Color del borde
            'text_wrap': True,
        })
    Formato_satis_import4 = workbook.add_format({
            'num_format': '0.0%',  # Formato de porcentaje con dos decimales
            'align': 'Center',          # Alinear el texto al centro
            'left': 1,                  # Borde izquierdo grueso
            'right': 1,                 # Borde derecho grueso
            'top': 1,                   # Borde superior delgado
            'bottom': 1,                # Borde inferior delgado
            'bg_color': '#FFFFFF',      # Color fondo gris
            'border_color': 'black',     # Color del borde
            'text_wrap': True,
        })
    Formato_satis_import5 = workbook.add_format({
            'num_format': '0.0%',  # Formato de porcentaje con dos decimales
            'align': 'Center',          # Alinear el texto al centro
            'left': 1,                  # Borde izquierdo grueso
            'right': 1,                 # Borde derecho grueso
            'top': 2,                   # Borde superior delgado
            'bottom': 2,                # Borde inferior delgado
            'bg_color': '#FFFFFF',      # Color fondo gris
            'border_color': 'black',     # Color del borde
            'text_wrap': True,
        })
    Formato_satis_import6 = workbook.add_format({
            'num_format': '0.00',  # Formato de porcentaje con dos decimales
            'align': 'Center',          # Alinear el texto al centro
            'left': 1,                  # Borde izquierdo grueso
            'right': 1,                 # Borde derecho grueso
            'top': 1,                   # Borde superior delgado
            'bottom': 1,                # Borde inferior delgado
            'bg_color': '#D3D3D3',      # Color fondo gris
            'border_color': 'black',     # Color del borde
            'text_wrap': True,
        })
    Formato_satis_import7 = workbook.add_format({
            'num_format': '0.00',  # Formato de porcentaje con dos decimales
            'align': 'Center',          # Alinear el texto al centro
            'left': 1,                  # Borde izquierdo grueso
            'right': 1,                 # Borde derecho grueso
            'top': 2,                   # Borde superior delgado
            'bottom': 2,                # Borde inferior delgado
            'bg_color': '#D3D3D3',      # Color fondo gris
            'border_color': 'black',     # Color del borde
            'text_wrap': True,
        })
    Formato_satis_import8 = workbook.add_format({
            'num_format': '0.0',  # Formato de porcentaje con dos decimales
            'align': 'Center',          # Alinear el texto al centro
            'left': 1,                  # Borde izquierdo grueso
            'right': 1,                 # Borde derecho grueso
            'top': 1,                   # Borde superior delgado
            'bottom': 1,                # Borde inferior delgado
            'bg_color': 'FFFFFF',      # Color fondo gris
            'border_color': 'black',     # Color del borde
            'text_wrap': True,
        })
    Formato_satis_import9 = workbook.add_format({
            'num_format': '0.0',  # Formato de porcentaje con dos decimales
            'align': 'Center',          # Alinear el texto al centro
            'left': 1,                  # Borde izquierdo grueso
            'right': 1,                 # Borde derecho grueso
            'top': 2,                   # Borde superior delgado
            'bottom': 2,                # Borde inferior delgado
            'bg_color': 'FFFFFF',      # Color fondo gris
            'border_color': 'black',     # Color del borde
            'text_wrap': True,
        })
    Formato_satis_import10 = workbook.add_format({
            'num_format': '0.00',  # Formato de porcentaje con dos decimales
            'align': 'Center',          # Alinear el texto al centro
            'left': 1,                  # Borde izquierdo grueso
            'right': 1,                 # Borde derecho grueso
            'top': 1,                   # Borde superior delgado
            'bottom': 1,                # Borde inferior delgado
            'bg_color': 'FFFFFF',      # Color fondo gris
            'border_color': 'black',     # Color del borde
            'text_wrap': True,
        })
    Formato_satis_import11 = workbook.add_format({
            'num_format': '0.0',  # Formato de porcentaje con dos decimales
            'align': 'Center',          # Alinear el texto al centro
            'left': 1,                  # Borde izquierdo grueso
            'right': 1,                 # Borde derecho grueso
            'top': 2,                   # Borde superior delgado
            'bottom': 2,                # Borde inferior delgado
            'bg_color': 'FFFFFF',      # Color fondo gris
            'border_color': 'black',     # Color del borde
            'text_wrap': True,
        })
    Formato_satis_import12 = workbook.add_format({
            'num_format': '0.0',  # Formato de porcentaje con dos decimales
            'align': 'Center',          # Alinear el texto al centro
            'left': 1,                  # Borde izquierdo grueso
            'right': 2,                 # Borde derecho grueso
            'top': 2,                   # Borde superior delgado
            'bottom': 2,                # Borde inferior delgado
            'bg_color': '#D3D3D3',      # Color fondo gris
            'border_color': 'black',     # Color del borde
            'text_wrap': True,
        })
    TG.merge_range(start_sati_impor+1, 1, start_sati_impor+2, 5, 'ATRIBUTO', Formato_satis_import1)
    TG.merge_range(start_sati_impor+1, 6, start_sati_impor+1, 7, 'NIP', Formato_satis_import1)
    TG.merge_range(start_sati_impor+1, 8, start_sati_impor+2, 8, 'P(NIP)', Formato_satis_import1)
    TG.merge_range(start_sati_impor+1, 9, start_sati_impor+1, 10, 'NSP', Formato_satis_import1)
    TG.merge_range(start_sati_impor+1,11, start_sati_impor+2, 11, 'P(NSP)', Formato_satis_import1)
    TG.merge_range(start_sati_impor+1,12, start_sati_impor+2, 12, 'CÓDIGO', Formato_satis_import1)
    TG.write(start_sati_impor+2,9,'0:100',Formato_satis_import1)
    TG.write(start_sati_impor+2,10,'0:5',Formato_satis_import1)
    TG.write(start_sati_impor+2, 6, 'CORR', Formato_satis_import1)
    TG.write(start_sati_impor+2, 7, 'PESO', Formato_satis_import1)
    TG.merge_range(start_sati_impor+len(Preguntas)+10,4, start_sati_impor+len(Preguntas)+10, 8, 'ISC Ponderado por NIP', workbook.add_format({'align': 'left', 'bg_color': '#C4D79B', 'left':2,'right':2,'top':2,'bottom':1,'border_color':'black'}))
    TG.merge_range(start_sati_impor+len(Preguntas)+11, 4,start_sati_impor+len(Preguntas)+11, 8, 'ISC Ponderado por NIP (1:5)', workbook.add_format({'align': 'left', 'bg_color': '#C4D79B', 'left':2,'right':2,'top':1,'bottom':2,'border_color':'black'}))

    def formatomatrizimportancia(worksheet,start_sati_impor,ite ):
        Formato_satis_import = workbook.add_format({
                'align': 'left',          # Alinear el texto al centro
                'left': 2,                  # Borde izquierdo grueso
                'right': 1,                 # Borde derecho grueso
                'top': 1,                   # Borde superior delgado
                'bottom': 1,                # Borde inferior delgado
                'bg_color': '#FFFFFF',      # Color fondo gris
                'border_color': 'black',     # Color del borde
                'text_wrap': True,
            })
        Formato_satis_import2 = workbook.add_format({
                'num_format': '0.00',  # Formato de porcentaje con dos decimales
                'align': 'Center',          # Alinear el texto al centro
                'left': 2,                  # Borde izquierdo grueso
                'right': 1,                 # Borde derecho grueso
                'top': 1,                   # Borde superior delgado
                'bottom': 1,                # Borde inferior delgado
                'bg_color': '#FFFFFF',      # Color fondo gris
                'border_color': 'black',     # Color del borde
                'text_wrap': True,
            })
        worksheet.merge_range(start_sati_impor+2, 1, start_sati_impor+2, 5, Preguntas[ite], Formato_satis_import)
        worksheet.write(start_sati_impor+2, 6, correl_visible(general, Preguntas[ite]), Formato_satis_import2)
    
    inicio = xl_rowcol_to_cell(start_sati_impor+3, 6)  
    for ite in range(len(Preguntas)):
        start_sati_impor=start_sati_impor+1
        formatomatrizimportancia(TG,start_sati_impor,ite)

    #Escribir la suma de correlacion
    fin = xl_rowcol_to_cell(start_sati_impor+2, 6)  
    TG.write_formula(start_sati_impor+3, 6, f'=IFERROR(SUM({inicio}:{fin}),0)', Formato_satis_import3)

    #Escribir los pesos
    inicio_peso = xl_rowcol_to_cell(id_row_peso+3, 7)  
    for ite in range(len(Preguntas)):
        id_row_peso=id_row_peso+1
        celda1 = xl_rowcol_to_cell(id_row_peso+2, 6)       
        celda2 = xl_rowcol_to_cell(start_sati_impor+3, 6, row_abs=True, col_abs=True) 
        TG.write(id_row_peso+2, 7, f'=IFERROR({celda1}/{celda2}, 0%)', Formato_satis_import4)

    #Calcular la media de los pesos
    fin_peso = xl_rowcol_to_cell(id_row_peso+2, 7)  
    TG.write_array_formula(id_row_peso+3, 7, id_row_peso+3, 7, f'=IFERROR(MEDIAN(IF({inicio_peso}:{fin_peso}<>0,{inicio_peso}:{fin_peso})),0)', Formato_satis_import5)


    Inicio_nip = xl_rowcol_to_cell(id__row_nip+3, 8)
    Inicio_nip_grafico = xl_rowcol_to_cell(id__row_nip+3, 8, row_abs=True, col_abs=True)
    for p in range(len(Preguntas)):
        id__row_nip=id__row_nip+1
        cel1 = xl_rowcol_to_cell(id__row_nip+2, 7)       
        cel2 = xl_rowcol_to_cell(start_sati_impor+3, 7, row_abs=True, col_abs=True) 
        TG.write(id__row_nip+2, 8, f'=IFERROR(({cel1}/{cel2})*100, 0)', Formato_satis_import6)
    #Escribir P(NIP)

    fin_nip = xl_rowcol_to_cell(id__row_nip+2, 8)
    fin_nip_grafico = xl_rowcol_to_cell(id__row_nip+2, 8, row_abs=True, col_abs=True)
    TG.write_array_formula(id__row_nip+3, 8, id__row_nip+3, 8, f'=IFERROR(AVERAGE(IF({Inicio_nip}:{fin_nip}<>0,{Inicio_nip}:{fin_nip})),0)', Formato_satis_import7)


    #ecribir el nsp de 0:100
    Inicio_nsp1= xl_rowcol_to_cell(id_row_nsp1+3, 9)
    for i in range(len(Preguntas)):
        id_row_nsp1=id_row_nsp1+1
        inicio, fin = Lista_rango_nsp[i].split(':')
        fila_ini, col_ini = xl_cell_to_rowcol(inicio) 
        fila_fin, col_fin = xl_cell_to_rowcol(fin)
        #Pesos correspondintes
        Pesos=[100,80,60,40,20]
        celdas = [xl_rowcol_to_cell(fila_ini, col) for col in range(col_ini,col_fin)]
        # Armar expresión de suma ponderada
        suma_ponderada = '+'.join([f'({celda}*{peso})' for celda, peso in zip(celdas, Pesos)])
        TG.write(id_row_nsp1+2, 9, f'=IFERROR({suma_ponderada}, 0)', Formato_satis_import8)
    fin_nsp1 = xl_rowcol_to_cell(id_row_nsp1+2, 9)
    TG.write_array_formula(id_row_nsp1+3, 9, id_row_nsp1+3, 9, f'=IFERROR(MEDIAN(IF({Inicio_nsp1}:{fin_nsp1}<>0,{Inicio_nsp1}:{fin_nsp1})),0)', Formato_satis_import9)

    #ecribir el nsp de 0:5
    Inicio_nsp2 = xl_rowcol_to_cell(id_row_nsp2+3, 10)
    for p in range(len(Preguntas)):
        id_row_nsp2=id_row_nsp2+1
        celd1 = xl_rowcol_to_cell(id_row_nsp2+2, 9)       
        TG.write(id_row_nsp2+2, 10, f'=IFERROR(({celd1}/20), 0)', Formato_satis_import10)
    #Escribir P(NIP)

    fin_nsp2 = xl_rowcol_to_cell(id_row_nsp2+2, 10)
    TG.write_array_formula(id_row_nsp2+3, 10, id_row_nsp2+3, 10, f'=IFERROR(MEDIAN(IF({Inicio_nsp2}:{fin_nsp2}<>0,{Inicio_nsp2}:{fin_nsp2})),0)', Formato_satis_import11)

    #Escribir el pnsp
    Inicio_pnsp = xl_rowcol_to_cell(id_row_pnsp+3,11)
    Inicio_pnsp_grafico = xl_rowcol_to_cell(id_row_pnsp+3,11, row_abs=True, col_abs=True)
    for p in range(len(Preguntas)):
        id_row_pnsp=id_row_pnsp+1
        cel1 = xl_rowcol_to_cell(id_row_pnsp+2, 10)       
        cel2 = xl_rowcol_to_cell(start_sati_impor+3, 10, row_abs=True, col_abs=True) 
        TG.write(id_row_pnsp+2, 11, f'=IFERROR(({cel1}/{cel2})*100, 0)', Formato_satis_import6)
    #Escribir P(NSP)

    fin_pnsp = xl_rowcol_to_cell(id__row_nip+2, 11)
    fin_pnsp_grafico = xl_rowcol_to_cell(id__row_nip+2, 11, row_abs=True, col_abs=True)
    TG.write_array_formula(id_row_pnsp+3, 11, id_row_pnsp+3, 11, f'=IFERROR(AVERAGE(IF({Inicio_pnsp}:{fin_pnsp}<>0,{Inicio_pnsp}:{fin_pnsp})),0)', Formato_satis_import12)

    #CODGIGO FORMATO CODICIONAL
    for c in range(len(Preguntas)):
        id_row_codigo += 1
        celdapnsp = xl_rowcol_to_cell(id_row_codigo + 2, 11)  # Columna L (11)
        celdapnip = xl_rowcol_to_cell(id_row_codigo + 2, 8)   # Columna I (8)

        # Formatos
        TG.write(id_row_codigo + 2, 12, 'V'+f'{c+1}', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border_color': 'black', 'left': 2, 'right': 2, 'top': 1, 'bottom': 1}))
        formato_rojo = workbook.add_format({'bg_color': '#FF0000', 'font_color': 'black','left': 2, 'right': 2,'valign': 'vcenter','top':1,'bottom':1,'align':'center','border_color': 'black'})
        formato_amarillo = workbook.add_format({'bg_color': '#FFFF00', 'font_color': 'black','left': 2, 'right': 2,'valign': 'vcenter','top':1,'bottom':1,'align':'center','border_color': 'black'})
        formato_verde = workbook.add_format({'bg_color': '#99CC00', 'font_color': 'black', 'left': 2, 'right': 2,'valign': 'vcenter','top':1,'bottom':1,'align':'center','border_color': 'black'})

        fila = id_row_codigo + 2
        columna = 12  # Columna M (12)

        # Regla 1: nip > 100 y nsp < 100
        TG.conditional_format(fila, columna, fila, columna, {
            'type': 'formula',
            'criteria': f'=AND({celdapnip}>100,{celdapnsp}<100)',
            'format': formato_rojo
        })

        # Regla 2: nip <= 100 y nsp < 100
        TG.conditional_format(fila, columna, fila, columna, {
            'type': 'formula',
            'criteria': f'=AND({celdapnip}<=100,{celdapnsp}<100)',
            'format': formato_amarillo
        })

        # Regla 3: nip >= 100 y nsp >= 100
        TG.conditional_format(fila, columna, fila, columna, {
            'type': 'formula',
            'criteria': f'=AND({celdapnip}>=100,{celdapnsp}>=100)',
            'format': formato_verde
        })


    TG.insert_image(id_row_grafica+3,14, 'fondo.png', {'x_scale': 0.767, 'y_scale': 0.847, 'x_offset': 69,'positioning': 1, 'object_position': 1})
    # Crear gráfico de dispersión
    chart = workbook.add_chart({'type': 'scatter', 'subtype': 'marker_only'})

    # Agregar serie
    chart.add_series({'name': 'Matriz de Importancia - Satisfacción',
        'categories': f"'T+G'!{Inicio_nip_grafico}:{fin_nip_grafico}" , # Columna A
        'values':     f"'T+G'!{Inicio_pnsp_grafico}:{fin_pnsp_grafico}",  # Columna B
        'marker':     {'type': 'diamond', 'size': 7, 'border': {'color': 'blue'}, 'fill': {'color': 'blue'}},
        'data_labels': {
            'value': True,  # Muestra el valor Y (satisfacción)
            'num_format': '0.00'  # Solo enteros
        }

    })

    # Quitar la leyenda
    chart.set_legend({'none': True})

    # Ejes con límites y pasos configurados, y líneas de cuadrícula activadas
    chart.set_x_axis({
        'name': 'Importancia',
        'min': 0,
        'max': 200,
        'num_format': '0',
        'major_unit': 20,
        'major_gridlines': {'visible': True,
                            'line': {'color': '#7F7F7F', 'width': 0.4, 'dash_type': 'solid'}
        }   
    })
    chart.set_y_axis({
        'name': 'Satisfacción',
        'min': 60,
        'max': 140,
        'num_format': '0',
        'major_unit': 5,
        'major_gridlines': {'visible': True,
                            'line': {'color': '#7F7F7F', 'width': 0.4, 'dash_type': 'solid'}
        }
    })

    # Fondo blanco simple
    chart.set_chartarea({
        'border': {'none': True},
        'fill': {'none': True}, 
    })
    chart.set_plotarea({
        'border':{'visible': True, 'color': 'black', 'width': 0.4, 'dash_type': 'solid'},
        'fill': {'none': True},
    })
    # Insertar gráfico
    TG.insert_chart(id_row_grafica,14, chart,{
            'x_scale': 1.212,
            'y_scale': 1.55
        })

    #Valor ISC ponderado por NIP
    TG.merge_range(id_row_isc+len(Preguntas)+10, 9, id_row_isc+len(Preguntas)+10, 10, f'=IFERROR(SUMPRODUCT({inicio_peso}:{fin_peso},{Inicio_nsp1}:{fin_nsp1}),0)', workbook.add_format({'align': 'center', 'bg_color': '#C4D79B', 'left':2,'right':2,'top':2,'bottom':1,'border_color':'black','num_format':'0.00'}))
    TG.merge_range(id_row_isc+len(Preguntas)+11, 9, id_row_isc+len(Preguntas)+11, 10, f'=IFERROR(({xl_rowcol_to_cell(id_row_isc+len(Preguntas)+10,9)}:{xl_rowcol_to_cell(id_row_isc+len(Preguntas)+10,10)})/20,0)', workbook.add_format({'align': 'center', 'bg_color': '#C4D79B', 'left':2,'right':2,'top':1,'bottom':2,'border_color':'black','num_format':'0.00'}))


#-------------------------------------------------metodologia---------------------------------------------------------------------------------
    #Crear la hoja de Metodología
    Metodologia = workbook.add_worksheet("Metodología")
    Metodologia.merge_range(0,0,1,11,'DESARROLLO DEL INDICE DE SERVICIO (INDICE DE SATISFACCIÓN)',workbook.add_format({'align': 'center', 'bold':True}))
    Metodologia.merge_range(2,0,2,11,'Cumplir con las expectativas del cliente (esto es, proporcionar los satisfactores) a menudo se considera como el mínimo requerido para conservar el negocio. ',workbook.add_format({'align': 'left'}))
    Metodologia.set_column('A:L', 11)
    Metodologia.merge_range(3,0,3,11,'Básicamente esto significa que el proveedor tiene que saber lo que los clientes quieren, debe clasificar estas necesidades bajo la forma de un índice de satisfacción y luego aplicar reingeniería para construir procesos y actividades que generen una lealtad genuina. ',workbook.add_format({'align': 'left','text_wrap':True}))
    Metodologia.set_row(3,30)
    Metodologia.merge_range(4,0,4,11,'A partir de los resultados cuantitativos se puede ilustrar en una gráfica bidimensional la relación entre el nivel de satisfacción reportado y la importancia de los procesos y atributos definidos dentro de la medición.',workbook.add_format({'align': 'left','text_wrap':True,'bold':True}))
    Metodologia.merge_range(5,0,5,11,'- La matriz resultante ofrece una visión sobre las acciones A seguir con el objetivo de mejorar la percepción del servicio por parte de los usuarios.',workbook.add_format({'align': 'left','text_wrap':True}))
    Metodologia.insert_image(
        9, 0, 'Matriz_importancia_satisfacion.png', {'x_scale': 0.913, 'y_scale': 0.9,'y_set':15, 'positioning': 1, 'object_position': 1}
    )
    Metodologia.merge_range(6,4,6,11,'La matriz importancia - satisfacción se puede dividir en cuadrantes que sugieren diferentes acciones a seguir:',workbook.add_format({'align': 'left','text_wrap':True,'bold':True}))
    Metodologia.merge_range(7,4,7,11,'- Los cuadrantes C1, C2, C3 y C4 constituyen el área crítica de mejoramiento y definen los atributos y procesos que requieren acciones novedosas para lograr el incremento en la percepción de la calidad por parte de los clientes, es decir es el área primaria de oportunidad donde es necesario invertir recursos; trabajar en ella potencializa un incremento importante en la satisfacción del cliente.',workbook.add_format({'align': 'left','text_wrap':True}))
    Metodologia.set_row(7,60)
    Metodologia.merge_range(8,4,8,11,'- Los cuadrantes M1, M2, M3 y M4 constituyen el área de mantenimiento e indican los atributos que están siendo satisfechos en la actualidad y que, dado su nivel de impacto en el NSC, no requieren acciones novedosas. Los cambios aquí tendrán poco impacto sobre la satisfacción y la lealtad; sin embargo, es necesario mantener estándares altos.',workbook.add_format({'align': 'left','text_wrap':True}))
    Metodologia.set_row(8,45)
    Metodologia.merge_range(9,4,9,11,'- Los cuadrantes I1, I2, I3 e I4 constituyen el área inerte y definen todos aquellos atributos y procesos que siendo poco satisfactorios no tienen el potencial para incidir de manera importante en la percepción general de la calidad del servicio ofrecido por la organización. Trabajar en ellos puede no conducir al mejoramiento sustancial en el NSC.',workbook.add_format({'align': 'left','text_wrap':True}))
    Metodologia.set_row(9,45)   
    Metodologia.merge_range(10,4,10,11,'- Los cuadrantes F3 y F4 representan el área de las fortalezas y define aquellos atributos sobre los cuales la organización debe sostener y dinamizar la imagen global y la percepción de calidad del servicio. Un trabajo constante y acciones novedosas son vitales para mantener estándares altos de calidad. Es importante resaltar el trabajo constante que se debe hacer sobre los cuadrantes F1 y F2 (área de mejora), al encontrarse muy cerca del área crítica de mejoramiento.',workbook.add_format({'align': 'left','text_wrap':True}))
    Metodologia.set_row(10,73)
    Metodologia.merge_range(11,4,11,11,'-Definición de las Variables de Medición',workbook.add_format({'align': 'left','text_wrap':True,'bold':True}))
    Metodologia.merge_range(12,4,12,11,'Para el cálculo del Nivel de Satisfacción del Cliente (Índice de Satisfacción) se tienen en cuenta dos variables: el Nivel de Importancia y el Nivel de Satisfacción (directamente proporcional al nivel de importancia) de cada uno de los atributos que conforman el proceso:',workbook.add_format({'align': 'left','text_wrap':True}))
    Metodologia.set_row(12,45)
    Metodologia.merge_range(13,4,13,11,'- El Nivel de Importancia permite determinar los perfiles de satisfacción (atributos con un nivel de importancia mayor a 60 puntos en Escala 0:100 ó 3 puntos en Escala 0:5), de cada proceso en particular, y el nivel de impacto del atributo en la satisfacción general; para esto, se calcula el nivel de importancia promedio NIP que tiene cada atributo en la satisfacción del cliente y el Nivel de Importancia que tiene cada proceso NIP* en la satisfacción general.',workbook.add_format({'align': 'left','text_wrap':True}))
    Metodologia.set_row(13,60)
    Metodologia.merge_range(14,4,14,11,'- Para determinar la satisfacción actual de los clientes encuestados, en cada categoría en particular, se calcula el Nivel de Satisfacción del Cliente NSC utilizando el procedimiento detallado en el Anexo G. Metodología Para el Cálculo del Nivel de Satisfacción del Cliente.',workbook.add_format({'align': 'left','text_wrap':True}))
    Metodologia.set_row(14,45)
    for fila in range(6,15):
        for columna in range(0,4):
            Metodologia.write(fila, columna, None, workbook.add_format({'align': 'center', 'bg_color': '#FFFFFF'}))
    Metodologia.merge_range(15,0,18,11,'METODOLOGÍA PARA EL CÁLCULO DEL INDICE DE SATISFACCIÓN DEL CLIENTE PONDERADO POR CORRELACIÓN',workbook.add_format({'align': 'left','valign':'vcenter','text_wrap':True,'left':2,'right':2,'top':2,'bottom':2,'border_color':'black','bold':True}))
    Metodologia.merge_range(19,0,19,11,'Calculo del nivel de Satisfacción General y Ponderado:',workbook.add_format({'align': 'left','valign':'vcenter','text_wrap':True,'bold':True,'left':2,'right':2,'top':2,'bottom':2,'border_color':'black'}))
    texto1 = (
        '1. Se realiza la conversión de las respuestas nominales a una escala numérica: cada calificación dada a un atributo, se convierte a una escala de 0 a 100, utilizando las siguientes equivalencias:\n'
        '5 - Excede notablemente las expectativas: 100 puntos\n'
        '4 - Supera las expectativas: 80 puntos\n'
        '3 – Cumple con las expectativas: 60 puntos\n'
        '2 – Por debajo de las expectativas: 40 puntos\n'
        '1 – Muy por debajo de las expectativas: 20 puntos'
    )
    Metodologia.merge_range(20,0,26,11,texto1,workbook.add_format({'align': 'left','valign':'vcenter','text_wrap':True,'left':2,'right':2,'top':1,'bottom':1,'border_color':'black'}))
    Metodologia.merge_range(27,0,27,11,'2. Se calcula el Nivel de Satisfacción General y por atributo con base en la siguiente fórmula:\n NSG / NSP  =  (n5 ×100)+ (n4 × 80)+ (n3 × 60)+ (n2 × 40)+ (n1 × 20)  /  (n5 + n4 + n3 + n2 + n1)',workbook.add_format({'align': 'left','valign':'vcenter','text_wrap':True,'left':2,'right':2,'top':1,'bottom':1,'border_color':'black'}))
    Metodologia.set_row(27,45)
    texto2 = (
        'Donde:\n'
        '    n5: número de clientes que calificaron Excede notablemente las expectativas\n'
        '    n4: número de clientes que calificaron Supera las expectativas\n'
        '    n3: número de clientes que calificaron Cumple con las expectativas\n'
        '    n2: número de clientes que calificaron Por debajo de las expectativas\n'
        '    n1: número de clientes que calificaron Muy por debajo de las expectativas'
    )
    Metodologia.merge_range(28,0,34,11,texto2,workbook.add_format({'align': 'left','valign':'vcenter','text_wrap':True,'left':2,'right':2,'top':1,'bottom':1,'border_color':'black'}))
    Metodologia.set_row(28,35)
    Metodologia.merge_range(35,0,35,11,'3. Se calcula el factor de ponderación P (NSP) correspondiente a los atributos de cada proceso, como la participación  de cada indice en la sumatoria.',workbook.add_format({'align': 'left','valign':'vcenter','text_wrap':True,'left':2,'right':2,'top':1,'bottom':1,'border_color':'black'}))
    Metodologia.set_row(35,30)
    texto3=('Para calcular el Nivel de Satisfacción del Cliente ponderado por Correlación se siguió la siguiente secuencia:\n'
            '1. Se calcula el indice de correlación de PEARSON entre la respuesta por atributo y la respuesta a la pregunta de satisfacción general.\n'
            '2. Se calcula la participación de cada correlación particular sobre la sumatoria de las mismas. \n'
            '3. Se calcula el indice de Satisfacción del cliente mediante la suma de los productos entre el peso que proviene de las correlaciones y el indice de satisfacción ponerado NSP. \n'
            '\n'
            'Donde:\n'
            'i: proceso de 1...m\n'
            'NSPi* : Nivel de satisfacción ponderado por NIP P: Ponderador (Participación en la sumatoria) \n'
            'NSP*i: Calificación del proceso i\n'
            )
    Metodologia.merge_range(36,0,49,11,texto3,workbook.add_format({'align': 'left','valign':'vcenter','text_wrap':True,'left':2,'right':2,'top':2,'bottom':2,'border_color':'black'}))
    Metodologia.insert_image(44,8,'formula.png', {'x_scale': 1.7, 'y_scale': 1.7,'y_set':15, 'positioning': 1, 'object_position': 1})
    if columnas_observaciones:
        for i, elemen in enumerate(columnas_observaciones, start=1):
            nombre_hoja = f'Comentarios u observaciones {i}'
            worksheet = workbook.add_worksheet(nombre_hoja)

            # Convertir a lista de Python para evitar mezclas con numpy
            datos = data[elemen].tolist()

            # Formato con wrap y alineación superior
            fmt = workbook.add_format({
                'align': 'left',
                'valign': 'top',
                'text_wrap': True,
            })

            # 1) Determinar ancho práctico (máximo 50 caracteres)
            max_len = max(len(str(x)) for x in [elemen] + datos if x is not None)
            wrap_width = min(max_len, 50)
            worksheet.set_column(0, 0, wrap_width, fmt)

            # 2) Escribir encabezado y ajustar altura según wrap
            header_paras = str(elemen).split('\n')
            header_lines = []
            for para in header_paras:
                header_lines.extend(textwrap.wrap(para, wrap_width) or [''])
            worksheet.write(0, 0, elemen, fmt)
            worksheet.set_row(0, len(header_lines) * 15)

            # 3) Escribir datos y ajustar altura de cada fila
            for row_idx, val in enumerate(datos, start=1):
                s = '' if val is None else str(val)
                paras = s.split('\n')
                wrapped_lines = []
                for para in paras:
                    wrapped_lines.extend(textwrap.wrap(para, wrap_width) or [''])
                worksheet.write(row_idx, 0, s, fmt)
                worksheet.set_row(row_idx, len(wrapped_lines) * 15)

            # 4) Crear la tabla
            worksheet.add_table(0, 0, len(datos), 0, {
                'columns': [{'header': elemen}],
                'name':    f'Tabla_{i}'
            })
    
    workbook.close()

    # ==============================================================
    # FILTROS DINÁMICOS — gestionados por fórmulas xlsxwriter
    # (Data Validation + cross-sheet _VISIBLE; sin VBA ni win32com)
    # ==============================================================
    if columnas_filtros_dinamicos and len(columnas_filtros_dinamicos) > 0:
        print(f"ℹ️ Filtros desplegables integrados para: {', '.join(columnas_filtros_dinamicos)}")
        sys.stdout.flush()
    else:
        print('ℹ️ No hay columnas de filtro seleccionadas.')
        sys.stdout.flush()
