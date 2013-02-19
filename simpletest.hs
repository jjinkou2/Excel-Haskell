{-# LANGUAGE OverloadedStrings, RecordWildCards #-}

import Data.Aeson
import qualified Data.Text as T 
import Data.Aeson.Encode.Pretty
import qualified Data.ByteString.Lazy.Char8 as BL

type Services = [ServStruct]


data ServStruct = ServStruct { servName :: T.Text , kpis :: [KpiStruct] } deriving (Show)

data KpiStruct = KpiStruct {kpiStructName :: T.Text, kpiStructData :: Value} deriving (Show)

instance ToJSON KpiStruct where
    toJSON KpiStruct{..} = object [kpiStructName .= kpiStructData]

instance ToJSON ServStruct where
    toJSON ServStruct{..} = object [servName .= kpis]


main = BL.putStrLn.servToBS $ services

servToBS :: Services -> BL.ByteString 
servToBS  = encode . map toJSON

services :: Services
services = map toServStruct servList

toServStruct :: String -> ServStruct
toServStruct s = ServStruct s (getKpitStruct s)

-- value test
servList :: [T.Text]
servList = ["BIV","BTIC","BIC"]
mos = toJSON ([4.3,2.3,4.9]::[Double])
comment = toJSON  (["Bon","Mauvais","Tres Bon"]::[String])

kpiTup = zip ["mos","CommentMos"] [mos,comment]
toKpiStruct dict = [KpiStruct x y |(x,y) <- dict ]

----test
getKpitStruct :: T.Text -> [KpiStruct]
getKpitStruct s = toKpiStruct kpiTup



----test

--kpisBIV = ServStruct "BIV"  (toKpiStruct kpiTup)
--kpisBIC = ServStruct "BIC" (toKpiStruct kpiTup)
--kpisBTIC = ServStruct "BTIC" (toKpiStruct kpiTup)
--services = [kpisBIV,kpisBIC, kpisBTIC]



