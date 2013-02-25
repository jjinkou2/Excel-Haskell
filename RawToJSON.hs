{-# LANGUAGE OverloadedStrings #-}
module RawToJSON where

import qualified Data.Map as M (fromList, lookup)
import Data.List.Split (endBy)

import Data.Aeson.Types (Pair,Value)

import Data.Aeson (toJSON,(.=),object)
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


kpis =  ["nbSites", "nbChannels", "nbCalls", "nbMinutes", "asr", "ner"
        ,"PGAD" ,"attps", "afis", "mos", "pdd", "csr", "rtd", "avail"
        , "unavail", "commentIndisp1", "commentIndisp2", "commentIndisp3"
        , "commentIndisp4", "commentIndisp5", "commentAFIS1", "commentAFIS2"
        ,  "commentMOS1", "commentMOS2"]
{--
    - tests 
        - --}
servToPair' :: [Value] -> T.Text -> Pair
servToPair' kpiValues s  = s .= kpisJSON
    where kpisJSON =  object $ zip kpis kpiValues


servToPair ::  [String] -> T.Text -> Pair
servToPair rows s  = s.= rawToVal rows




rawToVal :: [String] -> Value 
rawToVal rows  = object $ zip kpis kpiValues

    where

    -- KPistruct
        kpiValues = [ nbSitesVal, nbChannelsVal, nbMinutesVal, asrVal, nerVal
                    , attpsVal, afisVal, mosVal, pddVal, csrVal, rtdVal, availVal
                    , unavailVal, commentIndisp1Val, commentIndisp2Val, commentIndisp3Val
                    , commentIndisp4Val, commentIndisp5Val, commentAFIS1Val  
                    , commentAFIS2Val, commentMOS1Val, commentMOS2Val]
                    
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




