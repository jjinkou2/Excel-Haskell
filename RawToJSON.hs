{-# LANGUAGE OverloadedStrings #-}
module RawToJSON where

import qualified Data.Map as M (fromList, lookup)
import Data.List.Split (endBy)

import KpiStructure

import Data.Aeson (toJSON)
import qualified Data.Text as T 

{--
    - helpers
        - --} 

-- take 2 decimals

    -- reads double or put 0 
toDouble :: String -> Double
toDouble xs = case (reads.chgComma.endBy "," $ xs :: [(Double,String)] ) of
    [(d,s)] -> d
    _ -> 0
    where 
        chgComma [x,y] = x ++ "." ++ (take 2 y)
        chgComma xs = concat xs

toInt xs = case (reads xs :: [(Int,String)] ) of
    [(d,s)] -> d
    _ -> 0


{--
    - tests 
        - --}


toServStruct serv rows = ServStruct serv (rawToStruct rows) 

rawToStruct :: [String] -> [KpiStruct]    
rawToStruct rows =  [ nbSites, nbChannels, nbMinutes, asr, ner, attps, afis, mos
                    , pdd, csr, rtd, avail, unavail, commentIndisp1
                    , commentIndisp2, commentIndisp3, commentIndisp4
                    , commentIndisp5, commentAFIS1, commentAFIS2
                    , commentMOS1, commentMOS2
                    ]

    where
    -- KPistruct
        nbSites         = KpiStruct "nbSites" nbSitesVal  
        nbChannels      = KpiStruct "nbChannels" nbChannelsVal
        nbMinutes       = KpiStruct "nbMinutes" nbMinutesVal 
        asr             = KpiStruct "asrVal"      asrVal       
        ner             = KpiStruct "ner"             nerVal       
        attps           = KpiStruct "attps"           attpsVal     
        afis            = KpiStruct "afis"             afisVal      
        mos             = KpiStruct "mos"               mosVal       
        pdd             = KpiStruct "pdd"               pddVal       
        csr             = KpiStruct "csr"               csrVal       
        rtd             = KpiStruct "rtd"               rtdVal       
        avail           = KpiStruct "avail"           availVal     
        unavail         = KpiStruct "unavail"       unavailVal   

        commentIndisp1  = KpiStruct "commentIndisp1" commentIndisp1Val 
        commentIndisp2  = KpiStruct "commentIndisp2" commentIndisp2Val
        commentIndisp3  = KpiStruct "commentIndisp3" commentIndisp3Val
        commentIndisp4  = KpiStruct "commentIndisp4" commentIndisp4Val
        commentIndisp5  = KpiStruct "commentIndisp5" commentIndisp5Val
        commentAFIS1    = KpiStruct "commentAFIS1" commentAFIS1Val  
        commentAFIS2    = KpiStruct "commentAFIS2" commentAFIS2Val  
        commentMOS1     = KpiStruct "commentMOS1" commentMOS1Val   
        commentMOS2     = KpiStruct "commentMOS2" commentMOS2Val   

    -- JSON Values
        nbSitesVal      = toJSON$lookupData toInt 0 rowSite 
        nbChannelsVal   = toJSON$lookupData toInt 0 rowChannels
        nbMinutesVal    = toJSON$lookupData toDouble 0 rowMinutes
        asrVal          = toJSON$lookupData toDouble 0 rowAsr
        nerVal          = toJSON$lookupData toDouble 0 rowNer
        attpsVal        = toJSON$lookupData toDouble 0 rowAttps
        afisVal         = toJSON$lookupData toDouble 0 rowAfis
        mosVal          = toJSON$lookupData toDouble 0 rowMos
        pddVal          = toJSON$lookupData toDouble 0 rowPdd
        csrVal          = toJSON$lookupData toDouble 0 rowCsr
        rtdVal          = toJSON$lookupData toDouble 0 rowRtd
        availVal        = toJSON$lookupData toDouble 0 rowAvail
        unavailVal      = toJSON$lookupData toDouble 0 rowUnAvail
        
        commentIndisp1Val = toJSON$lookupData id "" rowComInd1
        commentIndisp2Val = toJSON$lookupData id ""  rowComInd2
        commentIndisp3Val = toJSON$lookupData id ""  rowComInd3
        commentIndisp4Val = toJSON$lookupData id ""  rowComInd4
        commentIndisp5Val = toJSON$lookupData id ""  rowComInd5
        commentAFIS1Val   = toJSON$lookupData id ""  rowComAfis1
        commentAFIS2Val   = toJSON$lookupData id ""  rowComAfis2
        commentMOS1Val    = toJSON$lookupData id ""  rowComMos1
        commentMOS2Val    = toJSON$lookupData id "" rowComMos2


        lookupData toX fill rowKpi = maybe (replicate 52 fill) -- default value [fill,...fill]
                                       (kpiData toX) -- handler 
                                       rowKpi -- Nothing or Just (kpi row) 
        rowSite     = rowNb "Sites"
        rowChannels = rowNb "Nb channels"
        rowMinutes  = rowNb "Minutes (Millions)"
        rowCalls    = rowNb "Calls (Millions)" 
        rowPgad     = rowNb "Post Gateway Answer Delay (sec)" 
        rowAsr      = rowNb "Answer Seizure Ratio (%)" 
        rowNer      = rowNb "Network Efficiency Ratio (%)" 
        rowAttps    = rowNb "ATTPS = Average Trouble Ticket Per Site" 
        rowAfis     = rowNb "Average FT Incident per Site\" AFIS" 
        rowMos      = rowNb "Mean Opinion Score (PESQ)" 
        rowPdd      = rowNb "Post Dialing Delay (sec)" 
        rowCsr      = rowNb "Call Sucessful Ratio" 
        rowRtd      = rowNb "RTD average" 
        rowAvail    = rowNb "Availability ratio HO (outage&changes)" 
        rowUnAvail  = rowNb "Unavailability minutes HO (outage&changes)" 
        rowComInd1  = rowNb "CommentIndispo1" 
        rowComInd2  = rowNb "CommentIndispo2" 
        rowComInd3  = rowNb "CommentIndispo3" 
        rowComInd4  = rowNb "CommentIndispo4" 
        rowComInd5  = rowNb "CommentIndispo5" 
        rowComAfis1 = rowNb "CommentAFIS1" 
        rowComAfis2 = rowNb "CommentAFIS2" 
        rowComMos1  = rowNb "CommentMOS1" 
        rowComMos2  = rowNb "CommentMOS2" 

        kpiData toX n  = map (toX.(rows!!).(+53*n)) [1..52]
        kpiName     = map (\n -> rows!!(53*n)) $ [0..72]
        kpiIndMap   = M.fromList $ zip kpiName [0..]
        rowNb s     = M.lookup s kpiIndMap


printListData :: [String] -> IO ()    
printListData rows = do     
    let kpiData toX n  = map (toX.(rows!!).(+53*n)) [1..52]
        kpiName     = map (\n -> rows!!(53*n)) $ [0..72]
        kpiIndMap   = M.fromList $ zip kpiName [0..]
        rowNb s     = M.lookup s kpiIndMap

        rowSite     = rowNb "Sites"
        rowChannels = rowNb "Nb channels"
        rowMinutes  = rowNb "Minutes (Millions)"
        rowCalls    = rowNb "Calls (Millions)" 
        rowPgad     = rowNb "Post Gateway Answer Delay (sec)" 
        rowAsr      = rowNb "Answer Seizure Ratio (%)" 
        rowNer      = rowNb "Network Efficiency Ratio (%)" 
        rowAttps    = rowNb "ATTPS = Average Trouble Ticket Per Site" 
        rowAfis     = rowNb "Average FT Incident per Site\" AFIS" 
        rowMos      = rowNb "Mean Opinion Score (PESQ)" 
        rowPdd      = rowNb "Post Dialing Delay (sec)" 
        rowCsr      = rowNb "Call Sucessful Ratio" 
        rowRtd      = rowNb "RTD average" 
        rowAvail    = rowNb "Availability ratio HO (outage&changes)" 
        rowUnAvail  = rowNb "Unavailability minutes HO (outage&changes)" 
        rowComInd1  = rowNb "CommentIndispo1" 
        rowComInd2  = rowNb "CommentIndispo2" 
        rowComInd3  = rowNb "CommentIndispo3" 
        rowComInd4  = rowNb "CommentIndispo4" 
        rowComInd5  = rowNb "CommentIndispo5" 
        rowComAfis1 = rowNb "CommentAFIS1" 
        rowComAfis2 = rowNb "CommentAFIS2" 
        rowComMos1  = rowNb "CommentMOS1" 
        rowComMos2  = rowNb "CommentMOS2" 

        lookupData toX fill rowKpi = maybe (replicate 52 fill) -- default value [fill,...fill]
                                       (kpiData toX) -- handler 
                                       rowKpi -- Nothing or Just (kpi row) 

        nbSites = lookupData toInt 0 rowSite 
        nbChannels = lookupData toInt 0 rowChannels
        nbMinutes = lookupData toDouble 0 rowMinutes
        asr = lookupData toDouble 0 rowAsr
        ner = lookupData toDouble 0 rowNer
        attps = lookupData toDouble 0 rowAttps
        afis = lookupData toDouble 0 rowAfis
        mos = lookupData toDouble 0 rowMos
        pdd = lookupData toDouble 0 rowPdd
        csr = lookupData toDouble 0 rowCsr
        rtd = lookupData toDouble 0 rowRtd
        avail = lookupData toDouble 0 rowAvail
        unavail = lookupData toDouble 0 rowUnAvail
        
        commentIndisp1 = lookupData id "" rowComInd1
        commentIndisp2 = lookupData id ""  rowComInd2
        commentIndisp3 = lookupData id ""  rowComInd3
        commentIndisp4 = lookupData id ""  rowComInd4
        commentIndisp5 = lookupData id ""  rowComInd5
        commentAFIS1 = lookupData id ""  rowComAfis1
        commentAFIS2 = lookupData id ""  rowComAfis2
        commentMOS1 = lookupData id ""  rowComMos1
        commentMOS2 = lookupData id "" rowComMos2

    
    -- return (kpiName!! ind , kpiData ind)
    -- print them
    putStrLn "----nbSites---"
    print nbSites
    putStrLn "----nbChannels---"
    print nbChannels
    putStrLn "----nbMinutes---"
    print nbMinutes
    putStrLn "----ASR ---"
    print asr
    putStrLn "----Ner ---"
    print ner
    putStrLn "----mos ---"
    print mos
    putStrLn "----pdd ---"
    print pdd
    putStrLn "----csr ---"
    print csr
    putStrLn "---- rtd ---"
    print rtd
    putStrLn "---- commentIndisp1 ---"
    print commentIndisp1
    putStrLn "---- commentIndisp2 ---"
    print commentIndisp2
    putStrLn "---- commentIndisp3 ---"
    print commentIndisp3
    putStrLn "---- commentIndisp4 ---"
    print commentIndisp4
    putStrLn "---- commentIndisp5 ---"
    print commentIndisp5
    putStrLn "---- commentAFIS1 ---"
    print commentAFIS1
    putStrLn "---- commentAFIS2 ---"
    print commentAFIS2
    putStrLn "---- commentMOS1 ---"
    print commentMOS1
    putStrLn "---- CommentMOS2 ---"
    mapM_ putStr commentMOS2
    putStrLn "----attps ---"
    print attps
    putStrLn "----afis ---"
    print afis 
    putStrLn "----avail ---"
    print avail 
    putStrLn "----unvail ---"
    print unavail 
