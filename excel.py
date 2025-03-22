import openpyxl
import subprocess
import time
import os

def ping_ips_from_excel():
    """
    Pide al usuario la ruta a un archivo de Excel y una letra de columna,
    luego hace ping a cada dirección IP encontrada en esa columna hasta que se encuentra una celda vacía.
    Las direcciones IP no alcanzables se guardan en el archivo "timeout.txt".
    """
    while True:
        excel_path = input("Introduce la ruta a tu archivo de Excel: ")
        try:
            workbook = openpyxl.load_workbook(excel_path)
            break
        except FileNotFoundError:
            print("Error: Archivo no encontrado. Por favor, introduce una ruta válida.")
        except Exception as e:
            print(f"Error al abrir el archivo de Excel: {e}")
            return

    while True:
        column_letter = input("Introduce la letra de la columna que contiene las direcciones IP (ej. A, B, C): ").upper()
        if not column_letter.isalpha() or len(column_letter) != 1:
            print("Entrada no válida. Por favor, introduce una sola letra para la columna.")
        else:
            break

    sheet = workbook.active  # Asumiendo que quieres usar la hoja activa
    timeout_file = "timeout.txt"

    if sheet.max_row == 0:
        print("La hoja de Excel está vacía.")
        return

    print("\nComenzando a hacer ping a las direcciones IP...")
    row_num = 1
    with open(timeout_file, 'w') as f_timeout:
        while True:
            cell = sheet[f"{column_letter}{row_num}"]
            ip_address = cell.value

            if ip_address is None or str(ip_address).strip() == "":
                print("\nSe alcanzó una celda vacía. Deteniendo el proceso.")
                break

            if not is_valid_ip(str(ip_address)):
                print(f"Saltando dirección IP no válida: {ip_address} en la fila {row_num}")
                row_num += 1
                continue

            print(f"\nHaciendo ping a la dirección IP: {ip_address} (Fila {row_num})")
            try:
                # Construye el comando ping según el sistema operativo
                if os.name == 'nt':  # Windows
                    command = ['ping', '-n', '1', '-w', '1000', str(ip_address)]
                else:  # Linux y macOS
                    command = ['ping', '-c', '1', '-W', '1', str(ip_address)]

                process = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                stdout, stderr = process.communicate(timeout=2) # Añadimos un timeout de 2 segundos
                return_code = process.returncode

                if return_code == 0:
                    print(f"  {ip_address} es alcanzable.")
                else:
                    print(f"  {ip_address} NO es alcanzable.")
                    f_timeout.write(f"{ip_address}\n")
                    if stderr:
                        print(f"  Error: {stderr.decode('utf-8').strip()}")

            except FileNotFoundError:
                print("Error: Comando 'ping' no encontrado en este sistema.")
                break
            except subprocess.TimeoutExpired:
                print(f"  Tiempo de espera agotado para {ip_address}. Guardado en {timeout_file}")
                f_timeout.write(f"{ip_address} (Timeout)\n")
            except Exception as e:
                print(f"Ocurrió un error al hacer ping a {ip_address}: {e}")

            time.sleep(1)  # Espera un breve período entre pings
            row_num += 1

def is_valid_ip(ip_str):
    """
    Comprueba si una cadena dada es una dirección IPv4 válida.
    """
    parts = ip_str.split('.')
    if len(parts) != 4:
        return False
    for part in parts:
        if not part.isdigit():
            return False
        if not 0 <= int(part) <= 255:
            return False
    return True

if __name__ == "__main__":
    import os
    ping_ips_from_excel()