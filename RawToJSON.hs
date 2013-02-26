{-# LANGUAGE OverloadedStrings #-}
module RawToJSON where

import qualified Data.Map as M (fromList, lookup)
import Data.List.Split (endBy)

import Data.Aeson.Types (Pair,Value)

import Data.Aeson (toJSON,(.=),object,encode)
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
servToPair :: [Value] -> T.Text -> Pair
servToPair kpiValues s  = s .= kpisJSON
    where kpisJSON =  object $ zip kpis kpiValues

--servToBS :: [Pair] -> BL.ByteString 
servToBS  = encode . object 
