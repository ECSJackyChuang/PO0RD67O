CREATE OR REPLACE PROCEDURE APPS.PO0RD67O (
    ERRBUF OUT VARCHAR2,RETCODE OUT NUMBER
    ,vSTART_DATE varchar2
    ,vPN varchar2)
    IS
    --2024/12 jacky added for lenovo sales team
    -- Define Sheet of Data --
    P_SHEETS     ECS_MRP_OOXML_PKG.SHEETS;
    -- Control Type for Construct pre-data --
    P_CONTROL    ECS_MRP_OOXML_PKG.R_CONTROL;
    -- Sheet output --
    X_OUTPUT    CLOB;
    X_OUTPUT_FULL    CLOB;
    -- Swap for speeding CLOB combine --
    X_SWAP        VARCHAR2(32000); 
    -- Rows and Cols --
    X_ROW NUMBER;
    X_COL NUMBER;
    -- Excel file name --
    P_FILE_NAME VARCHAR2(30):= 'LNV'||'_'||to_char(sysdate,'MMDDHHMiSS');--ECS_MRP_OOXML_PKG.C_DEFAULT_FNAME;
    -- Temp Diretory, stores excel files --
    P_DIR        VARCHAR2(30):= ECS_MRP_OOXML_PKG.C_DEFAULT_DIR;
    -- Generate Mode with Column generating or not --
    P_MODE         VARCHAR2(30):='N';
    -- Default Stylysh --
    P_STYLE ECS_MRP_OOXML_PKG.R_SETTING_STYLE;
    -- Stylish Setting --  
    P_SET_STYLE ECS_MRP_OOXML_PKG.R_SETTING_STYLE;
    FONTS         ECS_MRP_OOXML_PKG.STYLE_FONT;
    FILLS         ECS_MRP_OOXML_PKG.STYLE_FILL;
    BORDS         ECS_MRP_OOXML_PKG.R_STYLE_BORDER;
    NUMS        ECS_MRP_OOXML_PKG.R_STYLE_NUM;
    -- Condition Color Setting (Optional) --
    CONDI_COL_STY   ECS_MRP_OOXML_PKG.R_COLOR_STYLISH;
    CONDI_COL_REC   ECS_MRP_OOXML_PKG.R_COLOR_CONDITION;
    P_STYLE_NORMAL     NUMBER;
    P_STYLE_NORMAL_1    NUMBER;
    P_STYLE_DATE       NUMBER;
    P_STYLE_NUM       NUMBER;
    P_STYLE_TITLE      NUMBER;
    P_STYLE_supply number;
    P_STYLE_demand number;
    
    P_PROGRAM_NAME VARCHAR2(30) := 'OOXML';
    P_REQ_ID NUMBER  := FND_GLOBAL.CONC_REQUEST_ID;
    URL VARCHAR2(300);
    X_SERVER_ROUTE VARCHAR2(300);
    ----------------------------------------------------------
    TYPE TABLE_LIST IS TABLE OF VARCHAR2(50) INDEX BY PLS_INTEGER;
    T_TABLE TABLE_LIST;
    C_TABLE TABLE_LIST;
    
    TYPE demand_qty_array_type IS TABLE OF NUMBER INDEX BY PLS_INTEGER;
    demand_qty_array demand_qty_array_type;
    
    v_start_date date := trunc(to_date(vSTART_DATE),'iw');--轉換為wk
    v_pn varchar2(30) := replace(trim(vPN),' ','');--'SB21D13676';
    v_type_control varchar2(10);
    
    --for wk52 control
    v_current_date date;
    v_iso_year varchar2(10);
    v_iso_week varchar2(10);
    
    
    v_showing_week number:= 26;--showing 26 weeks
    v_week_diff number:=0;
    
    v_max_seq number;
    v_week_control varchar2(10):='iw';
    
    X_SHEET_CNT number;
    
    cursor cu_item is
    select distinct trim(CUS_PN) CUS_PN
    FROM ecs_lenovo_demand_pull_t
    WHERE CUS_PN = NVL(v_pn,CUS_PN)--p_pn
    and version_week >= to_number(to_char(v_start_date,'WW'))--排除當週以前的資料
    and disable_date is null
    order by CUS_PN;
    
    cursor cu_site(p_pn varchar2) is
    select distinct site,CUS_PN
    FROM ecs_lenovo_demand_pull_t
    WHERE trim(CUS_PN) = p_pn
    and version_week >= to_number(to_char(v_start_date,'WW'))--排除當週以前的資料
    and disable_date is null
    order by site; 
    
    cursor cu_version(p_pn varchar2, p_site varchar2) is
    select distinct wk_version
    ,version_year,version_week
    ,site,cus_pn,SOI
    ,last_pull_qty
    ,demand_qty_wk1, demand_qty_wk2, demand_qty_wk3, demand_qty_wk4, demand_qty_wk5
    ,demand_qty_wk6, demand_qty_wk7, demand_qty_wk8, demand_qty_wk9, demand_qty_wk10
    ,demand_qty_wk11, demand_qty_wk12, demand_qty_wk13, demand_qty_wk14, demand_qty_wk15
    ,demand_qty_wk16, demand_qty_wk17, demand_qty_wk18, demand_qty_wk19, demand_qty_wk20
    ,demand_qty_wk21, demand_qty_wk22, demand_qty_wk23, demand_qty_wk24, demand_qty_wk25
    ,demand_qty_wk26
    from ecs_lenovo_demand_pull_t xx
    where trim(CUS_PN) = p_pn
    and site = p_site
    and version_week >= to_number(to_char(v_start_date,'WW'))--排除當週以前的資料
    and disable_date is null
    order by WK_VERSION
    ;
begin

    FND_FILE.PUT_LINE(FND_FILE.LOG,'seq~:'||v_max_seq);
    dbms_output.put_line('seq:'||v_max_seq);
    
    -- Set Stylish 1 for sheet 1 (Optional) --
    FONTS.SZ := 10;
    FONTS.NAME := 'Calibri';
    -- Borders  (Optional) --
    BORDS.BORDER (1).NAME := 'left';
    BORDS.BORDER (1).STYLE := 'thin';
    BORDS.BORDER (2).NAME := 'right';
    BORDS.BORDER (2).STYLE := 'thin';
    BORDS.BORDER (3).NAME := 'top';
    BORDS.BORDER (3).STYLE := 'thin';
    BORDS.BORDER (4).NAME := 'bottom';
    BORDS.BORDER (4).STYLE := 'thin';
    -- Defined into Setting files  (Optional) --
    P_SET_STYLE.FONT := FONTS;
    P_SET_STYLE.FILL := FILLS;
    P_SET_STYLE.BORD := BORDS;
    -- Set cols width ( Optional ) --
    -- Set 1 - 5 cols to with 12 widths, 6-11 cols to 20 widths
    P_STYLE.COLS(1).WIDTH := 12;
    P_STYLE.COLS(1).SPAN := 6;
    
    ------------------------------------------------------------------------
    -- Register for Number format  (Optional) --
    ------------------------------------------------------------------------
    -- Define user defined type of number format
    NUMS.FMTID := 177;
    NUMS.FORMAT := '0.0000_ ';
    P_CONTROL.T_NUM(1) := NUMS;
    NUMS.FMTID := 178;
    NUMS.FORMAT := '0.00000%'; 
    P_CONTROL.T_NUM(2) := NUMS;
    NUMS.FMTID := 179;
    NUMS.FORMAT := 'm/d'; 
    P_CONTROL.T_NUM(3) := NUMS;
    -- You Should Call this procedure if you define color before generating sheets (SST)
    -- M$ pre-define first two type style in system, generate default for avoiding overwrite user define
    ECS_MRP_OOXML_PKG.GEN_DEFAULT_COLOR( P_STYLE , P_CONTROL );
    -- Regist Stylish for General Texts and Number --
    P_STYLE_NORMAL  := ECS_MRP_OOXML_PKG.GET_OOXML_STYCNT ( P_SET_STYLE,P_CONTROL,0 );
    P_STYLE_NORMAL_1 := ECS_MRP_OOXML_PKG.GET_OOXML_STYCNT ( P_SET_STYLE,P_CONTROL,0 );
    -- Regist Stylish for Date --
    P_STYLE_DATE    := ECS_MRP_OOXML_PKG.GET_OOXML_STYCNT ( P_SET_STYLE,P_CONTROL,179 );
    -- Regist Stylish for Number Format --
    P_STYLE_NUM        :=  ECS_MRP_OOXML_PKG.GET_OOXML_STYCNT ( P_SET_STYLE,P_CONTROL,177 );
    -- Fills Cols  (Optional) --
    FILLS.PATTERNTYPE := 'solid';
    FILLS.P_FGCOLOR := 'FFF0F0F0';
    P_SET_STYLE.FILL := FILLS;
    P_STYLE_TITLE   := ECS_MRP_OOXML_PKG.GET_OOXML_STYCNT ( P_SET_STYLE,P_CONTROL,0 );
    
    --color control    
    FILLS.P_FGCOLOR := 'FFFF37';--f2eee5
    P_SET_STYLE.FILL := FILLS;
    P_STYLE_demand   := ECS_MRP_OOXML_PKG.GET_OOXML_STYCNT ( P_SET_STYLE,P_CONTROL,0 );
    
    FILLS.PATTERNTYPE := 'solid';
    FILLS.P_FGCOLOR := '93FF93';
    P_SET_STYLE.FILL := FILLS;
    P_STYLE_supply   := ECS_MRP_OOXML_PKG.GET_OOXML_STYCNT ( P_SET_STYLE,P_CONTROL,0 );
    
    
    X_SHEET_CNT := 1;
    
    for lr_item in cu_item loop--loop for items
    
        for lr_site in cu_site(lr_item.CUS_PN) loop--loop for site
            
            X_OUTPUT := '';
            X_SWAP     := '';
            -- Sheet Initialize --
            ECS_MRP_OOXML_PKG.SET_SHEET(X_SHEET_CNT);
            -- Sheet Start --
            ECS_MRP_OOXML_PKG.ADD_OOXML_SST_AUTO( X_OUTPUT, X_SWAP, P_STYLE, P_CONTROL, X_ROW );
            
            -- Row start --
            ECS_MRP_OOXML_PKG.ADD_OOXML_ROW_AUTO( X_OUTPUT, X_SWAP, X_ROW, X_COL );
            
            ECS_MRP_OOXML_PKG.ADD_OOXML_COL_AUTO( X_OUTPUT, X_SWAP,X_ROW, X_COL , 'Site', P_STYLE_TITLE ,P_CONTROL );
            ECS_MRP_OOXML_PKG.ADD_OOXML_COL_AUTO( X_OUTPUT, X_SWAP,X_ROW, X_COL , 'Week', P_STYLE_TITLE ,P_CONTROL );
            ECS_MRP_OOXML_PKG.ADD_OOXML_COL_AUTO( X_OUTPUT, X_SWAP,X_ROW, X_COL , 'SOI', P_STYLE_TITLE ,P_CONTROL );
            
            /*印出26週需求的日期: ex:24wk40 */
            --前一週
            ECS_MRP_OOXML_PKG.ADD_OOXML_COL_AUTO( X_OUTPUT, X_SWAP,X_ROW, X_COL ,
                --TO_CHAR(trunc(trunc(trunc(v_start_date) -7,'iw')-1/86400,'iw'),'YY')
                TO_CHAR(trunc(trunc(trunc(v_start_date) -7,'iw'),'iw'),'YY')
                ||'WK'
                ||TO_CHAR(trunc(trunc(v_start_date) -7,'iw'),'WW') 
                , P_STYLE_TITLE ,P_CONTROL );
            
            for i in 0..(v_showing_week-1) loop
                
                --20241213 jacky added
                v_current_date := trunc(v_start_date + i*7, 'IW'); -- Ensure the date is Monday
                v_iso_year := to_char(v_current_date, 'YY');--IYYY
                v_iso_week := to_char(v_current_date, 'IW');
            
                -- Check if the week is 53 and adjust if it crosses the year boundary
                if v_iso_week = '53' and to_char(v_current_date, 'MM') = '01' then
                    v_iso_week := '01';
                    v_iso_year := to_char(v_current_date, 'YY');
                end if;
            
            
--                --Mon is default start day. 
--                ECS_MRP_OOXML_PKG.ADD_OOXML_COL_AUTO( X_OUTPUT, X_SWAP,X_ROW, X_COL ,
----                TO_CHAR(trunc(trunc(trunc(v_start_date) + I*7,'iw')-1/86400,'iw'),'YY')
--                TO_CHAR(trunc(trunc(trunc(v_start_date) + I*7,'iw'),'iw'),'YY')
--                ||'WK'
--                ||TO_CHAR(trunc(trunc(v_start_date) + I*7,'iw'),'WW') 
--                , P_STYLE_TITLE ,P_CONTROL );
                
--                ECS_MRP_OOXML_PKG.ADD_OOXML_COL_AUTO( X_OUTPUT, X_SWAP,X_ROW, X_COL ,
--                TO_CHAR(trunc(trunc(trunc(v_start_date) + I*7,'iw'),'iw'),'YY')
--                ||'WK'
--                ||TO_CHAR(trunc(trunc(v_start_date) + I*7,'iw'),'WW') 
--                , P_STYLE_TITLE ,P_CONTROL );
                
                ECS_MRP_OOXML_PKG.ADD_OOXML_COL_AUTO( X_OUTPUT, X_SWAP,X_ROW, X_COL ,
                v_iso_year || 'WK' ||v_iso_week
                , P_STYLE_TITLE ,P_CONTROL );
--                dbms_output.put_line(
--                    to_char(v_current_date, 'DD-MON-YY') || '~' ||
--                    v_iso_year || 'WK' ||
--                    v_iso_week
--                );
                
            end loop;
            -- Row end --
            ECS_MRP_OOXML_PKG.ADD_OOXML_RND_AUTO( X_OUTPUT, X_SWAP, X_ROW );
            
            -- Row start --
            ECS_MRP_OOXML_PKG.ADD_OOXML_ROW_AUTO( X_OUTPUT, X_SWAP, X_ROW, X_COL );
            
            --印出品名
            ECS_MRP_OOXML_PKG.ADD_OOXML_COL_AUTO( X_OUTPUT, X_SWAP,X_ROW, X_COL , lr_site.site, P_STYLE_NORMAL ,P_CONTROL );
            ECS_MRP_OOXML_PKG.ADD_OOXML_COL_AUTO( X_OUTPUT, X_SWAP,X_ROW, X_COL , lr_site.CUS_PN, P_STYLE_NORMAL ,P_CONTROL );
            ECS_MRP_OOXML_PKG.ADD_OOXML_COL_AUTO( X_OUTPUT, X_SWAP,X_ROW, X_COL , '', P_STYLE_NORMAL ,P_CONTROL );    

            --前一週
            ECS_MRP_OOXML_PKG.ADD_OOXML_COL_AUTO( X_OUTPUT, X_SWAP,X_ROW, X_COL , TO_CHAR(v_start_date-7,'MM/DD'), P_STYLE_TITLE ,P_CONTROL );
            --印出當週起始日:週一日期
            for i in 0..(v_showing_week-1) loop
                --Monday is default start day. 
                ECS_MRP_OOXML_PKG.ADD_OOXML_COL_AUTO( X_OUTPUT, X_SWAP,X_ROW, X_COL , TO_CHAR(v_start_date + I*7,'MM/DD'), P_STYLE_TITLE ,P_CONTROL );
            end loop;
            
            -- Row end --
            ECS_MRP_OOXML_PKG.ADD_OOXML_RND_AUTO( X_OUTPUT, X_SWAP, X_ROW );
            
            for lr_version in cu_version(lr_site.cus_pn,lr_site.site) loop
                
                v_week_diff :=0;
                v_week_diff := to_number(to_char(v_start_date,'WW')) - lr_version.version_week;
                
                -- 初始化 26 週的需求量到陣列中
                demand_qty_array(1) := lr_version.demand_qty_wk1;
                demand_qty_array(2) := lr_version.demand_qty_wk2;
                demand_qty_array(3) := lr_version.demand_qty_wk3;
                demand_qty_array(4) := lr_version.demand_qty_wk4;
                demand_qty_array(5) := lr_version.demand_qty_wk5;
                demand_qty_array(6) := lr_version.demand_qty_wk6;
                demand_qty_array(7) := lr_version.demand_qty_wk7;
                demand_qty_array(8) := lr_version.demand_qty_wk8;
                demand_qty_array(9) := lr_version.demand_qty_wk9;
                demand_qty_array(10) := lr_version.demand_qty_wk10;
                demand_qty_array(11) := lr_version.demand_qty_wk11;
                demand_qty_array(12) := lr_version.demand_qty_wk12;
                demand_qty_array(13) := lr_version.demand_qty_wk13;
                demand_qty_array(14) := lr_version.demand_qty_wk14;
                demand_qty_array(15) := lr_version.demand_qty_wk15;
                demand_qty_array(16) := lr_version.demand_qty_wk16;
                demand_qty_array(17) := lr_version.demand_qty_wk17;
                demand_qty_array(18) := lr_version.demand_qty_wk18;
                demand_qty_array(19) := lr_version.demand_qty_wk19;
                demand_qty_array(20) := lr_version.demand_qty_wk20;
                demand_qty_array(21) := lr_version.demand_qty_wk21;
                demand_qty_array(22) := lr_version.demand_qty_wk22;
                demand_qty_array(23) := lr_version.demand_qty_wk23;
                demand_qty_array(24) := lr_version.demand_qty_wk24;
                demand_qty_array(25) := lr_version.demand_qty_wk25;
                demand_qty_array(26) := lr_version.demand_qty_wk26;
                --row start
                ECS_MRP_OOXML_PKG.ADD_OOXML_ROW_AUTO( X_OUTPUT, X_SWAP, X_ROW, X_COL );
                ECS_MRP_OOXML_PKG.ADD_OOXML_COL_AUTO( X_OUTPUT, X_SWAP,X_ROW, X_COL , '26 wk '||lr_version.site, P_STYLE_NORMAL ,P_CONTROL );
                ECS_MRP_OOXML_PKG.ADD_OOXML_COL_AUTO( X_OUTPUT, X_SWAP,X_ROW, X_COL , lr_version.wk_version, P_STYLE_DATE ,P_CONTROL );
                ECS_MRP_OOXML_PKG.ADD_OOXML_COL_AUTO( X_OUTPUT, X_SWAP,X_ROW, X_COL , lr_version.SOI, P_STYLE_NORMAL ,P_CONTROL );
    --            ECS_MRP_OOXML_PKG.ADD_OOXML_COL_AUTO( X_OUTPUT, X_SWAP,X_ROW, X_COL , lr_version.demand_qty_wk1, P_STYLE_NORMAL ,P_CONTROL );
    --            ECS_MRP_OOXML_PKG.ADD_OOXML_COL_AUTO( X_OUTPUT, X_SWAP,X_ROW, X_COL , lr_version.demand_qty_wk2, P_STYLE_NORMAL ,P_CONTROL );
    --            ECS_MRP_OOXML_PKG.ADD_OOXML_COL_AUTO( X_OUTPUT, X_SWAP,X_ROW, X_COL , lr_version.demand_qty_wk3, P_STYLE_NORMAL ,P_CONTROL );
                
                /*-- 使用迴圈來輸出每週的需求量
                FOR i IN 1..26 LOOP
                    
                    if v_week_diff <> 0 then
                        if v_week_diff > 0 then--past
                            if (i-v_week_diff) < 0 then
    --                            if (v_week_diff-i) = 0 then--表示為前一週
    --                                ECS_MRP_OOXML_PKG.ADD_OOXML_COL_AUTO( X_OUTPUT, X_SWAP,X_ROW, X_COL , lr_version.LAST_PULL_QTY , P_STYLE_NORMAL ,P_CONTROL );
    --                            else
                                    null;--非前一週不印
    --                            end if;
                            else
                                ECS_MRP_OOXML_PKG.ADD_OOXML_COL_AUTO(
                                X_OUTPUT, X_SWAP, X_ROW, X_COL, demand_qty_array(i), P_STYLE_NORMAL, P_CONTROL);
                            end if;
                        else--v_week_diff < 0: future
                            if (i + v_week_diff) <= 1 then
                                if (i-ABS(v_week_diff)-1) = 0 then--表示前一週
                                    ECS_MRP_OOXML_PKG.ADD_OOXML_COL_AUTO( X_OUTPUT, X_SWAP,X_ROW, X_COL , lr_version.LAST_PULL_QTY , P_STYLE_NORMAL ,P_CONTROL );
                                else--非前一週不印
                                    ECS_MRP_OOXML_PKG.ADD_OOXML_COL_AUTO( X_OUTPUT, X_SWAP,X_ROW, X_COL , '*' , P_STYLE_NORMAL ,P_CONTROL );
                                end if;
                            else
                                ECS_MRP_OOXML_PKG.ADD_OOXML_COL_AUTO(
                                X_OUTPUT, X_SWAP, X_ROW, X_COL, demand_qty_array(i+v_week_diff), P_STYLE_NORMAL, P_CONTROL);
                            end if;
                        
                        end if;
                    else
                        --regular
                        if i=1 then
                            ECS_MRP_OOXML_PKG.ADD_OOXML_COL_AUTO(
                            X_OUTPUT, X_SWAP, X_ROW, X_COL, lr_version.LAST_PULL_QTY, P_STYLE_NORMAL, P_CONTROL);
                            
                            ECS_MRP_OOXML_PKG.ADD_OOXML_COL_AUTO(
                            X_OUTPUT, X_SWAP, X_ROW, X_COL, demand_qty_array(i), P_STYLE_NORMAL, P_CONTROL);
                        else
                            ECS_MRP_OOXML_PKG.ADD_OOXML_COL_AUTO(
                            X_OUTPUT, X_SWAP, X_ROW, X_COL, demand_qty_array(i), P_STYLE_NORMAL, P_CONTROL);    
                        end if;
                    end if;
                    
                END LOOP;*/
                if v_week_diff > 0 then --past   
                
                    for i in (1+v_week_diff)-1..26 loop--從起始週開始印
                        ECS_MRP_OOXML_PKG.ADD_OOXML_COL_AUTO(
                            X_OUTPUT, X_SWAP, X_ROW, X_COL, demand_qty_array(i), P_STYLE_demand, P_CONTROL);
                    end loop;
                    --補後段空白
                    for i in 1..v_week_diff loop
                        ECS_MRP_OOXML_PKG.ADD_OOXML_COL_AUTO(X_OUTPUT, X_SWAP, X_ROW, X_COL, '0', P_STYLE_demand, P_CONTROL);
                    end loop;
                    
                elsif v_week_diff = 0 then--current week
                    
                    ECS_MRP_OOXML_PKG.ADD_OOXML_COL_AUTO(
                            X_OUTPUT, X_SWAP, X_ROW, X_COL, lr_version.LAST_PULL_QTY, P_STYLE_supply, P_CONTROL);
                    for i in 1..26 loop
                        ECS_MRP_OOXML_PKG.ADD_OOXML_COL_AUTO(
                            X_OUTPUT, X_SWAP, X_ROW, X_COL, demand_qty_array(i), P_STYLE_demand, P_CONTROL);
                    end loop;
                elsif v_week_diff < 0 then --future
                    --補前段空白
                    for i in 1..(ABS(v_week_diff)+1) loop
                        if i <> (ABS(v_week_diff)+1) then                    
                            ECS_MRP_OOXML_PKG.ADD_OOXML_COL_AUTO(X_OUTPUT, X_SWAP, X_ROW, X_COL, '*', P_STYLE_supply, P_CONTROL);
                        else
                            ECS_MRP_OOXML_PKG.ADD_OOXML_COL_AUTO(X_OUTPUT, X_SWAP, X_ROW, X_COL, lr_version.LAST_PULL_QTY, P_STYLE_supply, P_CONTROL);
                        end if;
                    end loop;
                    
                    for i in 1..(26-ABS(v_week_diff)) loop
                        ECS_MRP_OOXML_PKG.ADD_OOXML_COL_AUTO(X_OUTPUT, X_SWAP, X_ROW, X_COL, demand_qty_array(i), P_STYLE_demand, P_CONTROL);
                    end loop;
                end if;
                
                -- Row end --
                ECS_MRP_OOXML_PKG.ADD_OOXML_RND_AUTO( X_OUTPUT, X_SWAP, X_ROW );
                
            end loop;
            
            

            -- Sheet END --
            --using X_SHEET_CNT control different sheet
            ECS_MRP_OOXML_PKG.ADD_OOXML_SND_WITH_COLOR( X_OUTPUT, X_SWAP, X_SHEET_CNT ,P_CONTROL ) ; 
            P_SHEETS(X_SHEET_CNT).SHEET_NAME := lr_site.SITE||'_'||lr_site.CUS_PN;
            P_SHEETS(X_SHEET_CNT).SHEET_TEXT := X_OUTPUT;

            X_SHEET_CNT := X_SHEET_CNT +1; 
            X_OUTPUT_FULL := X_OUTPUT_FULL||X_OUTPUT;
        
        end loop;--sheet control
    end loop;

--    ECS_MRP_OOXML_PKG.ADD_OOXML_SND_WITH_COLOR( X_OUTPUT, X_SWAP, 1,P_CONTROL ) ;
--    P_SHEETS(1).SHEET_NAME := 'PO0RD67O';
--    P_SHEETS(1).SHEET_TEXT := X_OUTPUT;
    
    ECS_MRP_DEBUG_PKG.REGIST('EXCEL');
    -- Generating Excel File --
    X_OUTPUT_FULL := ECS_MRP_OOXML_PKG.GEN_WORKBOOK_FIL( 
                            P_SHEETS,
                            P_FILE_NAME, 
                            P_DIR,
                            P_MODE,
                            P_STYLE,
                            P_CONTROL );
    ECS_MRP_DEBUG_PKG.R_END('EXCEL');
    -- Get URL of Excel Files --
    URL := ECS_MRP_OOXML_PKG.GET_FILE_INTO_URL ( P_DIR , P_FILE_NAME||'.xlsx' , P_PROGRAM_NAME , P_REQ_ID );
    -- Create OUTPUT redirect for Excel File --
    ECS_MRP_OOXML_PKG.CREATE_REDIRECT (URL);
    -- Get Directory route on server --
    SELECT DIRECTORY_PATH INTO X_SERVER_ROUTE FROM DBA_DIRECTORIES
    WHERE DIRECTORY_NAME = P_DIR;
    -- ######## BE CAREFULLY WHEN APPLYING THIS PROCESS #########
    -- Delete Temp excel file in Diretory --
    X_OUTPUT:= ecs_zip_files.DELETE_ALL (X_SERVER_ROUTE||'/' || P_FILE_NAME ||'.xlsx');
    DBMS_OUTPUT.PUT_LINE(URL);
    ECS_MRP_DEBUG_PKG.FLUSH;
    
--    IF tmpCount > 0 THEN
--            begin 
--            Apps.Ecs_Procsendemail(--寄信功能Procedure
--                L_Body,          -- P_TXT --內文：html_table列出的資料
--                NULL,            -- P_TXT2
--                NULL,            -- P_TXT3
--                NULL,            -- P_LOB
--                L_Subject,            -- P_SUB
--                v_to_mail,            -- P_RECEIVER     --用cursor串起來的收件人
--                v_cc_mail,            -- P_CC_RECEIVER  --用cursor串起來的副本收件人
--                'jacky.chuang@ecs.com.tw',            -- P_BCC_RECEIVER
--                NULL,            -- P_FILENAME
--                null,---593,            -- P_ALERT_ID
--                null,--'GROUP_1',     -- P_GROUP_NAME
--                0       -- P_TYPE
--                );
----            exception when others then
----                dbms_output.put_line('error:'||SQLCODE||'~'||SQLERRM);
--            end;
--        --fnd_file.put_line(fnd_file.log,'ERROR:'||SQLCODE||'~'||SQLERRM);
--        --dbms_output.put_line('mail sent.');
--        fnd_file.put_line(fnd_file.log,'mail sent');
--    else 
--        --dbms_output.put_line('no data.');
--        fnd_file.put_line(fnd_file.log,'no data.');
--    END IF;


end;
/
