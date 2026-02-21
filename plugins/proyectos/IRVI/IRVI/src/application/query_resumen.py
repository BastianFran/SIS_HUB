# -*- coding: utf-8 -*-
"""
Created on Thu Jun 26 08:38:15 2025

@author: CSANTOS
"""
import logging
import pandas as pd
def qry_resumen_cliente(connection, rut_completo, fecha_str):
    """
    Consulta los datos de resumen para un cliente específico y una fecha dada.
    """
    rut_truncado = rut_completo[:-1]  # Eliminar el último dígito
    sql = """
SELECT    
        sru.cnp_sru Fondo,
        A.tip_ope_ord Operacion, 
        A.ta_ser_nom_ser Instrumento, 
        SUM(b.cnt_asg) Cantidad,  
        SUM(ROUND(b.cnt_asg * b.pre_asg)) Monto
        --A.con_liq_ord Con_Liq, 
        --c.fec_trs Fecha,
        --to_char(c.fec_liq_trs,'dd-mm-yyyy')  Fecha_Pago 
        
FROM   
        ta_ult_trs c,  
        ta_obn A,  
        ta_asg b,
        tg_sru sru

WHERE   b.fec_asg >= TO_DATE(:1, 'DD-MM-YYYY')
        and a.tg_SRU_TG_CLI_TG_per_rut_per = :2
        AND  b.ta_obn_num_ord = A.num_ord  
        AND  b.ta_ult_trs_cor_trs = c.cor_trs
        and SRU.TG_CLI_TG_PER_RUT_PER = a.TG_SRU_TG_CLI_TG_PER_RUT_PER 
        and SRU.SRU_SRU = a.TG_SRU_SRU_SRU 


GROUP BY 
    sru.cnp_sru,
    A.tip_ope_ord,
    A.ta_ser_nom_ser
    --A.con_liq_ord,
    --c.fec_trs,
    --TO_CHAR(c.fec_liq_trs, 'dd-mm-yyyy')
    """

    try:
        resumen_df = pd.read_sql(sql, connection, params=(fecha_str, rut_truncado))
        print(f"\n Resultados de resumen para RUT {rut_truncado} y fecha {fecha_str}:")
        logging.info("Consulta de resumen ejecutada para RUT %s y fecha %s",
                     rut_truncado, fecha_str)

        if resumen_df.empty:
            print("No se encontraron registros de resumen.")
            logging.info("No se encontraron registros de resumen.")
        else:
            print(resumen_df.to_string(index=False))
            logging.info("Se encontraron %d registros de resumen.", len(resumen_df))

        return resumen_df

    except pd.io.sql.DatabaseError as db_error:
        print(f" Error de base de datos al ejecutar la consulta de resumen: {db_error}")
        logging.error("Error de base de datos al ejecutar la consulta de resumen: %s", db_error)
        return pd.DataFrame()

    except (TypeError, ValueError, RuntimeError) as error:
        print(f"Error inesperado al ejecutar la consulta de resumen: {error}")
        logging.error("Error inesperado al ejecutar la consulta de resumen: %s", error)
        return pd.DataFrame()


# SELECT 
#         sru.cnp_sru Fondo, 
#         decode(OBN.tip_ope_ord,'VTA','V','C') Operacion,  
#         OBN.ta_ser_nom_ser Instrumento,  
#         cg_fnc_gbl.fg_fmt_num(sum(ASG.cnt_asg),0,0) Cantidad, 
#         cg_fnc_gbl.fg_fmt_num(trunc(sum(ASG.cnt_asg * ASG.pre_asg)),0,0) Monto, 
#         to_char(asg.fec_asg,'dd-mm-yyyy') Fecha_trade, 
#         to_char(trs.fec_liq_trs,'dd-mm-yyyy')  Fecha_Pago 
            
# FROM
#         ta_obn obn, 
#         ta_asg asg, 
#         ta_ult_trs trs, 
#         tg_sru sru
         
# WHERE
#         obn.num_ord = asg.ta_obn_num_ord 
#         AND asg.fec_asg >= TO_DATE(:1, 'DD-MM-YYYY')
#         and obn.tg_SRU_TG_CLI_TG_per_rut_per = :2
#         and trs.cor_trs = asg.ta_ult_trs_cor_trs 
#         and SRU.TG_CLI_TG_PER_RUT_PER = OBN.TG_SRU_TG_CLI_TG_PER_RUT_PER 
#         and SRU.SRU_SRU = OBN.TG_SRU_SRU_SRU 
        
#     group by sru.cnp_sru, 
#              to_char(asg.fec_asg,'dd-mm-yyyy'), 
#              decode(OBN.tip_ope_ord,'VTA','V','C') , 
#              OBN.ta_ser_nom_ser, 
#              obn.tg_sru_sru_sru, 
#              to_char(trs.fec_liq_trs,'dd-mm-yyyy')