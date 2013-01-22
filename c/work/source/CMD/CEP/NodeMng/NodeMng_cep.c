LOCAL void  NodeMng_cep_convInsideNodeCode(void);

/* *********************************************
 *
 * ********************************************* */
EXPORT ER   NodeMng_cep_convInsideNodeCode(
                void
            )
{
    ER erRet;
    
    erRet = SAG_ctl_dec_sDes();
    erRet = NodeMng_cep_jdg_nodeCodeType();
    
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
LOCAL void  NodeMng_cep_jdg_nodeCodeType(
                void
            )
{
    ER erRet;
    
    erRet = MRN_ctl_set_rand();
    (void)BSP_COPY();
    
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
    
    erRet = Stg_cep_set_stg();
    
    return;
}
