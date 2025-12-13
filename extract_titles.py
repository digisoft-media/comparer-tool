from bs4 import BeautifulSoup
import os

def extract_titles_from_html(html_file_path):
    """
    Extrae todos los títulos (texto dentro de divs con class="ansiout")
    que se encuentran dentro de divs con data-testid="CommandResult"
    dentro del contenedor principal con data-testid="notebook-cell-output-container"
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
        # No necesitamos el contenedor principal, podemos buscar directamente
        command_results = soup.find_all('div', {'data-testid': 'CommandResult'})
        
        print(f"[OK] Se encontraron {len(command_results)} elementos CommandResult")
        
        # Lista para almacenar los títulos encontrados
        titles = []
        
        # Iterar sobre cada CommandResult
        for idx, command_result in enumerate(command_results, 1):
            try:
                found_title = False
                
                # ESTRATEGIA 1: Buscar div con data-testid="ansi-output" y luego div con class="ansiout"
                # Esta es la estructura real: CommandResult > ... > div[data-testid="ansi-output"] > div.ansiout
                ansi_output_divs = command_result.find_all('div', {'data-testid': 'ansi-output'})
                
                if ansi_output_divs:
                    for ansi_output in ansi_output_divs:
                        # Buscar div con class="ansiout" dentro del ansi-output
                        # Puede tener solo la clase "ansiout" o múltiples clases
                        ansiout_divs = ansi_output.find_all('div', class_=lambda x: x and 'ansiout' in x)
                        
                        for ansiout_div in ansiout_divs:
                            text = ansiout_div.get_text(strip=True)
                            if text and text not in titles:  # Evitar duplicados
                                titles.append(text)
                                found_title = True
                                print(f"\n[{idx}] [OK] Título encontrado:")
                                print(f"    {text}")
                
                # ESTRATEGIA 2: Buscar directamente div con class que contenga "ansiout" en cualquier lugar del CommandResult
                # Usar lambda para buscar divs cuya clase contenga "ansiout" (puede tener múltiples clases)
                if not found_title:
                    ansiout_divs = command_result.find_all('div', class_=lambda x: x and 'ansiout' in x)
                    
                    for ansiout_div in ansiout_divs:
                        text = ansiout_div.get_text(strip=True)
                        if text and text not in titles:  # Evitar duplicados
                            titles.append(text)
                            found_title = True
                            print(f"\n[{idx}] [OK] Título encontrado (método 2 - ansiout directo):")
                            print(f"    {text}")
                
                # ESTRATEGIA 4: Buscar por texto que contenga "User Selects" como fallback
                if not found_title:
                    # Buscar cualquier div que contenga el texto "User Selects"
                    all_divs = command_result.find_all('div', string=lambda text: text and 'User Selects' in text)
                    for div in all_divs:
                        text = div.get_text(strip=True)
                        if text:
                            titles.append(text)
                            found_title = True
                            print(f"\n[{idx}] [OK] Título encontrado (método 4 - búsqueda por texto 'User Selects'):")
                            print(f"    {text}")
                            break
                
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
        print(f"Total de títulos encontrados: {len(titles)}")
        print(f"{'='*60}\n")
        
        # Imprimir todos los títulos
        if titles:
            print("LISTA COMPLETA DE TÍTULOS:")
            print("-" * 60)
            for i, title in enumerate(titles, 1):
                print(f"{i}. {title}")
        else:
            print("No se encontraron títulos.")
        
        return titles
        
    except FileNotFoundError:
        print(f"Error: No se encontró el archivo {html_file_path}")
        return []
    except Exception as e:
        print(f"Error durante la ejecución: {e}")
        import traceback
        traceback.print_exc()
        return []

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
    
    # Extraer los títulos
    titles = extract_titles_from_html(html_file)
