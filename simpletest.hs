{-# LANGUAGE OverloadedStrings, RecordWildCards #-}

import Data.Aeson
import qualified Data.Text as T 
import Data.Aeson.Encode.Pretty
import qualified Data.ByteString.Lazy.Char8 as BL

import KpiStructure


instance ToJSON KpiStruct where
    toJSON KpiStruct{..} = object [kpiName .= kpiData]

instance ToJSON ServStruct where
    toJSON ServStruct{..} = object [servName .= kpis]


main = BL.putStrLn.servToBS $ services

servToBS :: Services -> BL.ByteString 
servToBS  = encode . map toJSON

services :: Services
services = map toServStruct servList

toServStruct :: T.Text -> ServStruct
toServStruct s = ServStruct s (getKpitStruct s)

-- value test
servList :: [T.Text]
servList = ["BIV","BTIC","BIC"]
mos = toJSON ([4.3,2.3,4.9]::[Double])
comment = rowDataToValue  (["Bon","Mauvais","Tres Bon"]::[String])

kpiTup = zip ["mos","CommentMos"] [mos,comment]
toKpiStruct dict = [KpiStruct x y |(x,y) <- dict ]

----test
getKpitStruct :: T.Text -> [KpiStruct]
getKpitStruct s = toKpiStruct kpiTup


rowDataToValue :: ToJSON a => a -> Value 
rowDataToValue = toJSON

----test

--kpisBIV = ServStruct "BIV"  (toKpiStruct kpiTup)
--kpisBIC = ServStruct "BIC" (toKpiStruct kpiTup)
--kpisBTIC = ServStruct "BTIC" (toKpiStruct kpiTup)
--services = [kpisBIV,kpisBIC, kpisBTIC]



