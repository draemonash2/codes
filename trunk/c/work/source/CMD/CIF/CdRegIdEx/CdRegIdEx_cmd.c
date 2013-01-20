LOCAL void  CdRegIdEx_cmd_set_seqInfo(void);
LOCAL void  CdRegIdEx_cmd_prtecStg(void);

/* *********************************************
 *
 * ********************************************* */
EXPORT ER   CdRegIdEx_cmd_chk_cmdPkt(
                void
            )
{
    ER erRet;
    
    erRet = NodeMng_cep_convInsideNodeCode();
    
    switch (erRet) {
        case E_OK:      /* None */  break;
        case E_NOSPT:   /* None */  break;
        default:
            (void)vact_trp(CMD_EXP, LOG_CODE_NODEMNG, (VP_INT)erRet);
    }
    return erRet;
}

/* *********************************************
 *
 * ********************************************* */
LOCAL void  CdRegIdEx_cmd_set_seqInfo(
                void
            )
{
    ER erRet;
    
    erRet = NodeMng_cep_convInsideNodeCode();
    (void)NodeMng_cep_jdg_nodeCodeType();
    
    return;
}

/* *********************************************
 *
 * ********************************************* */
LOCAL void  CdRegIdEx_cmd_prtecStg(
                void
            )
{
    ER erRet;
    
    erRet = NodeMng_cep_convInsideNodeCodeForEndSrv();
    
    AthEnc_cep_convInsideNodeCode();
    
    return;
}
