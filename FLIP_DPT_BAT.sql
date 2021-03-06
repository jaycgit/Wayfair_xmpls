#"flips" specific groups of "batches" that have been exported in the past 10 minutes and re-enables them for export to another system.

UPDATE    P_BAT
SET              PB_S_STATUS = 0 
FROM         I_M INNER JOIN
                      P_ITEMS ON I_M.ITEM_ID = P_ITEMS.ITEM_ID INNER JOIN
                      P_BAT ON P_ITEMS.B_NUMBER = P_BAT.B_NUMBER INNER JOIN
                      ITM_S_FLGS ON I_M.ITEM_ID = ITM_S_FLGS.ITEM_ID
WHERE     (P_BAT.PB_S_STATUS = 4) AND (I_M.STORE_P_DEP IN ('03')) AND 
                      (P_BAT.REC_STAT_DTE > DATEADD(mi, -10, CURRENT_TIMESTAMP)) AND (P_BAT.STORE_ID IN ('4','5','6','7')) 
--AND (ITM_S_FLGS.SMR_S_LABEL_2 >= 1)
