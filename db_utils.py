import oracledb

def conectar_bd():
    print("\n=== Conectando a Base de Datos Oracle (Modo Grueso) ===")
    try:
        oracledb.init_oracle_client(lib_dir=r"C:\oracle\instantclient_23_0")
        conn = oracledb.connect(
            user="datasoft",
            password="data2001",
            dsn="172.16.14.17:1521/ORCL"
        )
        print("Conexión exitosa a la BD.")
        return conn
    except Exception as e:
        print(f"Error al conectar a BD: {e}")
        return None

def ejecutar_consulta_db(query, conn):
    if not conn:
        print("Aviso: No hay conexión a BD para ejecutar la consulta.")
        return "0"
    try:
        with conn.cursor() as cursor:
            cursor.execute(query)
            res = cursor.fetchone()
            if res:
                return str(res[0])
            return "0"
    except Exception as e:
        print(f"Error ejecutando consulta en BD: {e}")
        return "0"

def ejecutar_actualizacion_db(query, conn):
    if not conn:
        print("Aviso: No hay conexión a BD para ejecutar la actualización.")
        return False
    try:
        with conn.cursor() as cursor:
            cursor.execute(query)
            conn.commit()
            return True
    except Exception as e:
        print(f"Error ejecutando actualización en BD: {e}")
        return False

def ejecutar_consulta_fila_db(query, conn):
    """Retorna la primera fila completa como dict o None."""
    if not conn:
        return None
    try:
        with conn.cursor() as cursor:
            cursor.execute(query)
            row = cursor.fetchone()
            if row and cursor.description:
                cols = [col[0] for col in cursor.description]
                return dict(zip(cols, row))
            return None
    except Exception as e:
        print(f"Error ejecutando consulta de fila en BD: {e}")
        return None
