'''
Consulta QUERY para las facturas imprimidas por soporte.
'''
import logging
import pandas as pd

def qry_folios_fecha(connection, rut_completo, fecha_str):
    '''Consulta SQL para cliente en ejecución.'''
    rut_truncado = rut_completo[:-1]
    sql = """
    SELECT  TG_SRU_TG_CLI_TG_PER_RUT_PER AS Rut,
            TG_SRU_SRU_SRU AS Sub_Rut,
            TO_CHAR(FEC_FAC, 'DD-MM-YYYY') AS Fecha_Operacion,
            FOL_FAC AS Folio
    FROM    ta_fac
    WHERE   FEC_FAC = TO_DATE(:1, 'DD-MM-YYYY')
            AND FOL_FAC IS NOT NULL
            AND TG_SRU_TG_CLI_TG_PER_RUT_PER = :2
            AND USR_IMP_FAC IN ('JCHANDIA','JSEPULVE',
                                'MCARVACH1',
                                'BBRUNA','CSANTOS')
    """
    try:
        resultados_df = pd.read_sql(sql, connection, params=(fecha_str, rut_truncado))
        print(f"\nConsulta ejecutada para RUT {rut_truncado} y fecha {fecha_str}")
        logging.info("Consulta ejecutada para RUT %s y fecha %s", rut_truncado, fecha_str)

        if resultados_df.empty:
            print("No se encontraron registros.")
            logging.info("No se encontraron registros.")
        else:
            print(resultados_df.to_string(index=False))
            logging.info("Se encontraron %d registros.", len(resultados_df))

        return resultados_df

    except pd.io.sql.DatabaseError as db_error:
        print(f"Error de base de datos al ejecutar la consulta: {db_error}")
        logging.error("Error de base de datos al ejecutar la consulta: %s", db_error)
        return pd.DataFrame()

    except (TypeError, ValueError, RuntimeError) as error:
        print(f"Error inesperado al ejecutar la consulta: {error}")
        logging.error("Error inesperado al ejecutar la consulta: %s", error)
        return pd.DataFrame()
    