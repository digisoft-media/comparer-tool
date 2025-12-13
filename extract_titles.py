from bs4 import BeautifulSoup
import os
from openpyxl import Workbook
from datetime import datetime

def extract_table_data(command_result_with_table):
    """
    Extrae los datos de la tabla dentro de un div CommandResult con tabla.
    Retorna una lista de diccionarios con los datos de la tabla.
    """
    table_data = []
    
    try:
        # Buscar la tabla con data-testid="datagrid.table"
        datagrid_table = command_result_with_table.find('div', {'data-testid': 'datagrid.table'})
        if not datagrid_table:
            # Intentar buscar por role="table"
            datagrid_table = command_result_with_table.find('div', {'role': 'table'})
        
        if not datagrid_table:
            return table_data
        
        # Buscar los encabezados de las columnas
        headers = []
        column_headers = datagrid_table.find_all('div', {'role': 'columnheader'})
        for header in column_headers:
            # Intentar obtener el nombre de la columna desde aria-label o data-cell-id
            aria_label = header.get('aria-label', '')
            cell_id = header.get('data-cell-id', '')
            
            # Buscar el texto del encabezado
            header_text = header.find('span', class_=lambda x: x and 'dg--header-text' in x)
            if header_text:
                header_name = header_text.get_text(strip=True)
            elif aria_label:
                header_name = aria_label
            elif cell_id:
                header_name = cell_id
            else:
                continue
            
            # Solo incluir las columnas que nos interesan (account, Avg Sales, Number of IDs)
            if header_name and header_name not in ['#row_number#']:
                headers.append(header_name)
        
        # Si no encontramos encabezados por los métodos anteriores, buscar directamente en los spans
        if not headers:
            # Buscar todos los headers por el texto
            header_spans = datagrid_table.find_all('span', class_=lambda x: x and 'dg--header-text' in x)
            headers = [span.get_text(strip=True) for span in header_spans if span.get_text(strip=True)]
        
        if not headers:
            return table_data
        
        # Buscar todas las filas de datos
        # Las celdas tienen role="cell" y contienen spans con class="dg--default-cell"
        rows = datagrid_table.find_all('div', {'role': 'row'})
        
        for row in rows:
            # Saltar la fila de encabezado
            if row.find('div', {'role': 'columnheader'}):
                continue
            
            cells = row.find_all('div', {'role': 'cell'})
            if not cells:
                continue
            
            row_data = {}
            for cell in cells:
                # Saltar las celdas de número de fila (tienen clase dg--row-number-cell)
                if 'dg--row-number-cell' in ' '.join(cell.get('class', [])):
                    continue
                
                # Buscar el valor en el span con class="dg--default-cell"
                value_span = cell.find('span', class_=lambda x: x and 'dg--default-cell' in x)
                if value_span:
                    value = value_span.get_text(strip=True)
                    
                    # Mapear la celda a la columna correcta usando data-cell-id
                    cell_id = cell.get('data-cell-id', '')
                    if cell_id:
                        # El formato es "0_account", "0_Avg Sales", etc.
                        parts = cell_id.split('_', 1)
                        if len(parts) == 2:
                            col_name = parts[1]
                            # Solo incluir si la columna está en los headers
                            if col_name in headers:
                                row_data[col_name] = value
            
            # Solo agregar la fila si tiene datos
            if row_data and len(row_data) > 0:
                table_data.append(row_data)
        
    except Exception as e:
        print(f"    [WARN] Error al extraer datos de tabla: {e}")
    
    return table_data

def extract_titles_from_html(html_file_path):
    """
    Extrae todos los títulos (texto dentro de divs con class="ansiout")
    que se encuentran dentro de divs con data-testid="CommandResult"
    y también extrae las tablas de datos del div hermano siguiente.
    Retorna una lista de diccionarios con 'title' y 'table_data'.
    """
    
    try:
        print(f"Leyendo archivo: {html_file_path}")
        
        # Leer el archivo HTML
        with open(html_file_path, 'r', encoding='utf-8') as f:
            html_content = f.read()
        
        # Parsear el HTML con BeautifulSoup
        soup = BeautifulSoup(html_content, 'html.parser')
        print("[OK] Archivo HTML parseado correctamente")
        
        # Buscar todos los divs con data-testid="CommandResult" en todo el documento
        command_results = soup.find_all('div', {'data-testid': 'CommandResult'})
        
        print(f"[OK] Se encontraron {len(command_results)} elementos CommandResult")
        
        # Lista para almacenar los resultados (título + tabla)
        results = []
        
        # Iterar sobre cada CommandResult
        for idx, command_result in enumerate(command_results, 1):
            try:
                found_title = False
                title_text = None
                
                # ESTRATEGIA 1: Buscar div con data-testid="ansi-output" y luego div con class="ansiout"
                ansi_output_divs = command_result.find_all('div', {'data-testid': 'ansi-output'})
                
                if ansi_output_divs:
                    for ansi_output in ansi_output_divs:
                        ansiout_divs = ansi_output.find_all('div', class_=lambda x: x and 'ansiout' in x)
                        for ansiout_div in ansiout_divs:
                            text = ansiout_div.get_text(strip=True)
                            if text:
                                title_text = text
                                found_title = True
                                print(f"\n[{idx}] [OK] Título encontrado:")
                                print(f"    {text}")
                                break
                        if found_title:
                            break
                
                # ESTRATEGIA 2: Buscar directamente div con class que contenga "ansiout"
                if not found_title:
                    ansiout_divs = command_result.find_all('div', class_=lambda x: x and 'ansiout' in x)
                    for ansiout_div in ansiout_divs:
                        text = ansiout_div.get_text(strip=True)
                        if text:
                            title_text = text
                            found_title = True
                            print(f"\n[{idx}] [OK] Título encontrado (método 2):")
                            print(f"    {text}")
                            break
                
                # ESTRATEGIA 3: Buscar por texto que contenga "User Selects"
                if not found_title:
                    all_divs = command_result.find_all('div', string=lambda text: text and 'User Selects' in text)
                    for div in all_divs:
                        text = div.get_text(strip=True)
                        if text:
                            title_text = text
                            found_title = True
                            print(f"\n[{idx}] [OK] Título encontrado (método 3):")
                            print(f"    {text}")
                            break
                
                # Si encontramos un título, buscar el div hermano siguiente con la tabla
                if found_title and title_text:
                    table_data = []
                    
                    # Buscar el siguiente div hermano con data-testid="CommandResult" y la clase específica
                    next_sibling = command_result.find_next_sibling('div', {'data-testid': 'CommandResult'})
                    
                    if next_sibling:
                        # Verificar si tiene la clase "command-result-tabs"
                        classes = next_sibling.get('class', [])
                        if 'command-result-tabs' in ' '.join(classes):
                            print(f"    [OK] Div con tabla encontrado, extrayendo datos...")
                            table_data = extract_table_data(next_sibling)
                            if table_data:
                                print(f"    [OK] {len(table_data)} filas extraídas de la tabla")
                            else:
                                print(f"    [WARN] No se encontraron datos en la tabla")
                    
                    # Agregar el resultado (título + tabla)
                    results.append({
                        'title': title_text,
                        'table_data': table_data
                    })
                
                # Si no se encontró nada, mostrar información de debug solo para los primeros 3
                if not found_title and idx <= 3:
                    print(f"\n[{idx}] [WARN] No se encontró título. Debug:")
                    try:
                        # Verificar estructura
                        ansi_output_count = len(command_result.find_all('div', {'data-testid': 'ansi-output'}))
                        ansiout_count = len(command_result.find_all('div', class_='ansiout'))
                        all_divs_count = len(command_result.find_all('div'))
                        print(f"    - Contenedores ansi-output: {ansi_output_count}")
                        print(f"    - Divs con class ansiout: {ansiout_count}")
                        print(f"    - Total de divs: {all_divs_count}")
                        
                        # Mostrar algunas clases de divs para debug
                        sample_divs = command_result.find_all('div')[:5]
                        print(f"    - Primeras 5 clases de divs encontradas:")
                        for i, div in enumerate(sample_divs, 1):
                            div_class = div.get('class', '(sin clase)')
                            div_text = div.get_text(strip=True)[:30] if div.get_text(strip=True) else '(vacío)'
                            print(f"      {i}. class='{div_class}' | text='{div_text}...'")
                    except Exception as e:
                        print(f"    Error en debug: {e}")
                        
            except Exception as e:
                print(f"\n[{idx}] [ERROR] Error al procesar CommandResult: {e}")
                import traceback
                traceback.print_exc()
                continue
        
        print(f"\n{'='*60}")
        print(f"Total de títulos encontrados: {len(results)}")
        print(f"{'='*60}\n")
        
        return results
        
    except FileNotFoundError:
        print(f"Error: No se encontró el archivo {html_file_path}")
        return []
    except Exception as e:
        print(f"Error durante la ejecución: {e}")
        import traceback
        traceback.print_exc()
        return []

def save_titles_to_excel(results, output_file="títulos.xlsx"):
    """
    Guarda los títulos y sus tablas asociadas en un archivo Excel.
    Cada título va seguido de su tabla de datos con las columnas: account, Avg Sales, Number of IDs
    """
    try:
        # Crear un nuevo workbook
        wb = Workbook()
        
        # Seleccionar la hoja activa (primera hoja)
        ws = wb.active
        ws.title = "Títulos"
        
        current_row = 1
        
        for result in results:
            title = result['title']
            table_data = result['table_data']
            
            # Escribir el título en negrita
            ws[f'A{current_row}'] = title
            ws[f'A{current_row}'].font = ws[f'A{current_row}'].font.copy(bold=True)
            current_row += 1
            
            # Si hay datos de tabla, escribir la tabla
            if table_data:
                # Obtener las columnas de la primera fila
                columns = list(table_data[0].keys())
                
                # Escribir encabezados de la tabla
                for col_idx, col_name in enumerate(columns, start=1):
                    cell = ws.cell(row=current_row, column=col_idx)
                    cell.value = col_name
                    cell.font = cell.font.copy(bold=True)
                current_row += 1
                
                # Escribir los datos de la tabla
                for row_data in table_data:
                    for col_idx, col_name in enumerate(columns, start=1):
                        value = row_data.get(col_name, '')
                        # Intentar convertir a número si es posible
                        try:
                            # Si parece un número, convertir a int o float
                            if isinstance(value, str) and value.replace(',', '').replace('.', '').isdigit():
                                value = float(value.replace(',', ''))
                        except:
                            pass
                        ws.cell(row=current_row, column=col_idx).value = value
                    current_row += 1
            
            # Agregar una fila en blanco entre títulos
            current_row += 1
        
        # Ajustar anchos de columnas
        ws.column_dimensions['A'].width = 50
        # Buscar todas las columnas usadas para ajustar sus anchos
        max_col = 1
        for result in results:
            if result['table_data']:
                max_col = max(max_col, len(result['table_data'][0].keys()))
        
        # Ajustar anchos de las columnas de datos (B en adelante)
        from openpyxl.utils import get_column_letter
        for col_idx in range(2, max_col + 2):
            col_letter = get_column_letter(col_idx)
            ws.column_dimensions[col_letter].width = 20
        
        # Guardar el archivo
        wb.save(output_file)
        print(f"[OK] Archivo Excel creado exitosamente: {output_file}")
        print(f"Total de títulos guardados: {len(results)}")
        return True
        
    except Exception as e:
        print(f"[ERROR] Error al crear el archivo Excel: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    # Lista de archivos HTML a probar en orden de prioridad
    html_files = [
        "body.html",
        "Dimension Scenarios.html",
        "Dimension Scenarios v3.html"
    ]
    
    html_file = None
    for file in html_files:
        if os.path.exists(file):
            html_file = file
            break
    
    if not html_file:
        print(f"Error: No se encontró ningún archivo HTML.")
        print("Archivos buscados:")
        for file in html_files:
            print(f"  - {file}")
        print("Por favor, asegúrate de que al menos uno de estos archivos existe en el directorio actual.")
        exit(1)
    
    print(f"Procesando archivo: {html_file}")
    print("=" * 60)
    
    # Extraer los títulos y tablas
    results = extract_titles_from_html(html_file)
    
    # Guardar los títulos y tablas en un archivo Excel
    if results:
        # Generar nombre de archivo con timestamp para evitar sobrescribir
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        excel_filename = f"títulos_{timestamp}.xlsx"
        save_titles_to_excel(results, excel_filename)
    else:
        print("\n[INFO] No se encontraron títulos para guardar en Excel.")
