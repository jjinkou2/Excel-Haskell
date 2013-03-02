{-# LANGUAGE OverloadedStrings, BangPatterns #-}

-- ghc --make XL2JSON.hs -o testBS -rtsopts -prof -auto-all -caf-all

import ExcelCom 
import qualified Data.Map as M (fromList, lookup)
import Data.List.Split (endBy)
import qualified Data.Text as T 
import qualified Data.Text.Read as T 
import Data.Aeson.Types (Pair,Value)
import Data.Aeson
import Data.Aeson.Encode.Pretty (encodePretty)
import qualified Data.ByteString.Lazy.Char8 as BL 
import Control.Concurrent.ParallelIO.Local
import qualified Data.Vector as V


fichierTest = "qos.xls"
fichierTest3 = "E:/Programmation/haskell/Com/qos.xls"
fichierTest4 = "C:/Users/lyup7323/Developpement/Haskell/Com/qos.xls"

sheetsName = ["BIV","BIC","BTIP_H323","BTIC","MCCE","OPITML","FIA","BTIP_SIP" ,"OVP","BTELU"]
--sheetsName = ["BIV","BIC","BTIP_H323","BTIC","MCCE","OPITML","BTIP_SIP" ,"OVP"]
--sheetsName = ["BIV","BIC","BTIP_H323","BTIC","MCCE","OPITML"]
--sheetsName = ["FIA"]

--sheetsName = ["MCCE","OPITML","FIA","BTIP_SIP"]
-- sheetsName = ["BIV","BIC","BTIP_H323","BTIC","MCCE","FIA"]
main = xl2json fichierTest3 >>= BL.writeFile "json.txt"
getData' sheet cast row = fmap V.fromList $ mapM (castText cast sheet row) [4..55] 
           
lookupData' s c f r = maybe (return $ V.replicate 52 f) (getData' s c) r

castText cast sheet row col = do 
    vals <- sheet # getCells row col ## getFormula 
    return.cast.T.pack $ vals

xl2json :: String -> IO BL.ByteString
xl2json file = coRun $ do 
    (pExl, workBooks, workSheets) <- xlInit file
    xs <- mapM (processRowData workSheets) sheetsName
   -- xs <- withPool 2 $ \pool -> parallel pool $ map (processRowData workSheets) sheetsName
    xlQuit workBooks pExl
    --stopGlobalPool
    return $ servToBS xs

xlQuit workBooks appXl = do
    workBooks # method_1_0 "Close" xlDoNotSaveChanges
    appXl # method_0_0 "Quit"

xlInit file = do   
    pExl <- createObjExl
    workBooks <- pExl # getWorkbooks
    pExl # propertySet "DisplayAlerts" [inBool False]
    workBook <- workBooks # openWorkBooks file
    putStrLn  $"File loaded: " ++ file
    workSheets <- workBook # getWSheets
    return (pExl, workBooks, workSheets)
   
debug :: String -> String -> String -> IO (IDispatch a) 
debug file sheetName rngStr = do
    (pExl, workBooks, workSheets) <- xlInit file
    sheetSel <- workSheets # propertyGet_1 "Item" sheetName
    rng <- sheetSel # propertyGet_1 "Range" rngStr

    return rng


processRowData :: Sheet a -> String -> IO Pair
processRowData sheets sheetName = do 
    putStrLn $ "got all datas from " ++ sheetName
    putStrLn $ replicate 50 '-'
    kpisVal <- valuesFromSheet sheets sheetName
    return $ servToPair kpisVal (T.pack sheetName)    

valuesFromSheet' :: Sheet a -> String -> IO [Value] 
valuesFromSheet' workSheets sheetName= do 
    sheetSel <- workSheets # propertyGet_1 "Item" sheetName
    rng <- sheetSel # propertyGet_1 "Range"  ("C7:C85"::String)
    (_,kpiNames) <- rng # enumVariants' "" :: IO (Int,[String]) 


    -- get KPI
    let kpiIndMap   = M.fromList $ zip kpiNames [7..]
        rowSite     = M.lookup "Sites" kpiIndMap
        rowChannels = M.lookup "Nb channels" kpiIndMap
        rowMinutes  = M.lookup "Minutes (Millions)" kpiIndMap
        rowCalls    = M.lookup "Calls (Millions)" kpiIndMap
        rowPgad     = M.lookup "Post Gateway Answer Delay (sec)" kpiIndMap
        rowAsr      = M.lookup "Answer Seizure Ratio (%)" kpiIndMap
        rowNer      = M.lookup "Network Efficiency Ratio (%)" kpiIndMap
        rowAttps    = M.lookup "ATTPS = Average Trouble Ticket Per Site" kpiIndMap
        rowAfis     = M.lookup "Average FT Incident per Site\" AFIS" kpiIndMap
        rowMos      = M.lookup "Mean Opinion Score (PESQ)" kpiIndMap
        rowPdd      = M.lookup "Post Dialing Delay (sec)" kpiIndMap
        rowCsr      = M.lookup "Call Sucessful Ratio" kpiIndMap
        rowRtd      = M.lookup "RTD average" kpiIndMap
        rowAvail    = M.lookup "Availability ratio HO (outage&changes)" kpiIndMap
        rowUnAvail  = M.lookup "Unavailability minutes HO (outage&changes)" kpiIndMap
        rowComInd1  = M.lookup "CommentIndispo1" kpiIndMap
        rowComInd2  = M.lookup "CommentIndispo2" kpiIndMap
        rowComInd3  = M.lookup "CommentIndispo3" kpiIndMap
        rowComInd4  = M.lookup "CommentIndispo4" kpiIndMap
        rowComInd5  = M.lookup "CommentIndispo5" kpiIndMap
        rowComAfis1 = M.lookup "CommentAFIS1" kpiIndMap
        rowComAfis2 = M.lookup "CommentAFIS2" kpiIndMap
        rowComMos1  = M.lookup "CommentMOS1" kpiIndMap
        rowComMos2  = M.lookup "CommentMOS2" kpiIndMap

        lookupData :: Variant a => a -> Maybe Int -> IO (V.Vector a) 
        -- return a list of Kpi's datas if kpi's row is found otherwise a defaut value
        lookupData fill rowKpi  = maybe (return $ V.replicate 52 fill) -- default value [fill,...fill]
                                       (getData fill) -- handler 
                                       rowKpi -- Nothing or Just (kpi's row) 

        getData :: Variant a => a -> Int -> IO (V.Vector a)
        getData def row = do
            rng <- sheetSel # propertyGet_1 "Range" (rngStr row)
            (_,kpiDatas) <- rng # enumVariants' def 
            return.V.fromList $ kpiDatas
            where
                rngStr row =  "D" ++ (show row) ++ ":BC" ++ (show row)


   -- print kpiIndMap
    nbSitesVal         <- fmap toJSON $ lookupData (0::Int) rowSite
    nbChannelsVal      <- fmap toJSON $ lookupData (0::Int) rowChannels
    nbMinutesVal       <- fmap (toJSON.V.map trunc) $ lookupData 0 rowMinutes
    nbCallsVal         <- fmap (toJSON.V.map trunc) $ lookupData 0 rowCalls
    postGADVal         <- fmap (toJSON.V.map trunc) $ lookupData 0 rowPgad
    asrVal             <- fmap (toJSON.V.map trunc) $ lookupData 0 rowAsr
    nerVal             <- fmap (toJSON.V.map trunc) $ lookupData 0 rowNer
    attpsVal           <- fmap (toJSON.V.map trunc) $ lookupData 0 rowAttps
    afisVal            <- fmap (toJSON.V.map trunc) $ lookupData 0 rowAfis 
    mosVal             <- fmap (toJSON.V.map trunc) $ lookupData 0 rowMos 
    pddVal             <- fmap (toJSON.V.map trunc) $ lookupData 0 rowPdd
    csrVal             <- fmap (toJSON.V.map trunc) $ lookupData 0 rowCsr 
    rtdVal             <- fmap (toJSON.V.map trunc) $ lookupData 0 rowRtd 
    availVal           <- fmap (toJSON.V.map trunc) $ lookupData 0 rowAvail 
    unavailVal         <- fmap (toJSON.V.map trunc) $ lookupData 0 rowUnAvail 
    commentIndisp1Val  <- fmap toJSON $ lookupData (""::String) rowComInd1 
    commentIndisp2Val  <- fmap toJSON $ lookupData (""::String) rowComInd2
    commentIndisp3Val  <- fmap toJSON $ lookupData (""::String) rowComInd3
    commentIndisp4Val  <- fmap toJSON $ lookupData (""::String) rowComInd4
    commentIndisp5Val  <- fmap toJSON $ lookupData (""::String) rowComInd5
    commentAFIS1Val    <- fmap toJSON $ lookupData (""::String) rowComAfis1
    commentAFIS2Val    <- fmap toJSON $ lookupData (""::String) rowComAfis2
    commentMOS1Val     <- fmap toJSON $ lookupData (""::String) rowComMos1
    commentMOS2Val     <- fmap toJSON $ lookupData (""::String) rowComMos2

 -- KPistruct
    let kpiValues = [ nbSitesVal, nbChannelsVal, nbCallsVal, nbMinutesVal, asrVal
                    , nerVal, postGADVal, attpsVal, afisVal, mosVal, pddVal
                    , csrVal, rtdVal, availVal, unavailVal, commentIndisp1Val
                    , commentIndisp2Val, commentIndisp3Val, commentIndisp4Val
                    , commentIndisp5Val, commentAFIS1Val  , commentAFIS2Val
                    , commentMOS1Val, commentMOS2Val]
    
    return kpiValues

    
valuesFromSheet :: Sheet a -> String -> IO [Value] 
valuesFromSheet workSheets sheetName= do 
    sheetSel <- workSheets # propertyGet_1 "Item" sheetName
    kpiNames <- mapM (\x -> sheetSel # getCells x 3 ## getText) [7..85]

    -- get KPI
    let kpiIndMap   = M.fromList $ zip kpiNames [7..]
        rowSite     = M.lookup "Sites" kpiIndMap
        rowChannels = M.lookup "Nb channels" kpiIndMap
        rowMinutes  = M.lookup "Minutes (Millions)" kpiIndMap
        rowCalls    = M.lookup "Calls (Millions)" kpiIndMap
        rowPgad     = M.lookup "Post Gateway Answer Delay (sec)" kpiIndMap
        rowAsr      = M.lookup "Answer Seizure Ratio (%)" kpiIndMap
        rowNer      = M.lookup "Network Efficiency Ratio (%)" kpiIndMap
        rowAttps    = M.lookup "ATTPS = Average Trouble Ticket Per Site" kpiIndMap
        rowAfis     = M.lookup "Average FT Incident per Site\" AFIS" kpiIndMap
        rowMos      = M.lookup "Mean Opinion Score (PESQ)" kpiIndMap
        rowPdd      = M.lookup "Post Dialing Delay (sec)" kpiIndMap
        rowCsr      = M.lookup "Call Sucessful Ratio" kpiIndMap
        rowRtd      = M.lookup "RTD average" kpiIndMap
        rowAvail    = M.lookup "Availability ratio HO (outage&changes)" kpiIndMap
        rowUnAvail  = M.lookup "Unavailability minutes HO (outage&changes)" kpiIndMap
        rowComInd1  = M.lookup "CommentIndispo1" kpiIndMap
        rowComInd2  = M.lookup "CommentIndispo2" kpiIndMap
        rowComInd3  = M.lookup "CommentIndispo3" kpiIndMap
        rowComInd4  = M.lookup "CommentIndispo4" kpiIndMap
        rowComInd5  = M.lookup "CommentIndispo5" kpiIndMap
        rowComAfis1 = M.lookup "CommentAFIS1" kpiIndMap
        rowComAfis2 = M.lookup "CommentAFIS2" kpiIndMap
        rowComMos1  = M.lookup "CommentMOS1" kpiIndMap
        rowComMos2  = M.lookup "CommentMOS2" kpiIndMap

        --lookupData :: Variant a => a -> Maybe Int -> IO [a] 
        -- return a list of Kpi's datas if kpi's row is found otherwise a defaut value
        lookupData cast fill rowKpi  = maybe (return $ V.replicate 52 fill) -- default value [fill,...fill]
                                       (getData cast) -- handler 
                                       rowKpi -- Nothing or Just (kpi's row) 

        getData cast row = fmap V.fromList $ mapM (castText cast sheetSel row) [4..55] 
            where
                castText cast sheet row col = getOneData sheet row col >>= 
                                              return.cast.T.pack
                getOneData sheet row col =  sheet # getCells row col ## getFormula

   -- print kpiIndMap
    nbSitesVal         <- fmap toJSON $ lookupData toInt 0 rowSite
    nbChannelsVal      <- fmap toJSON $ lookupData toInt 0 rowChannels
    nbMinutesVal       <- fmap toJSON $ lookupData toDouble 0.0 rowMinutes
    nbCallsVal         <- fmap toJSON $ lookupData toDouble 0.0 rowCalls
    postGADVal         <- fmap toJSON $ lookupData toDouble 0.0 rowPgad
    asrVal             <- fmap toJSON $ lookupData toDouble 0.0 rowAsr
    nerVal             <- fmap toJSON $ lookupData toDouble 0.0 rowNer
    attpsVal           <- fmap toJSON $ lookupData toDouble 0.0 rowAttps
    afisVal            <- fmap toJSON $ lookupData toDouble 0.0 rowAfis 
    mosVal             <- fmap toJSON $ lookupData toDouble 0.0 rowMos 
    pddVal             <- fmap toJSON $ lookupData toDouble 0.0 rowPdd
    csrVal             <- fmap toJSON $ lookupData toDouble 0.0 rowCsr 
    rtdVal             <- fmap toJSON $ lookupData toDouble 0.0 rowRtd 
    availVal           <- fmap toJSON $ lookupData toDouble 0.0 rowAvail 
    unavailVal         <- fmap toJSON $ lookupData toDouble 0.0 rowUnAvail 
    commentIndisp1Val  <- fmap toJSON $ lookupData id (T.pack "") rowComInd1 
    commentIndisp2Val  <- fmap toJSON $ lookupData id (T.pack "") rowComInd2
    commentIndisp3Val  <- fmap toJSON $ lookupData id (T.pack "") rowComInd3
    commentIndisp4Val  <- fmap toJSON $ lookupData id (T.pack "") rowComInd4
    commentIndisp5Val  <- fmap toJSON $ lookupData id (T.pack "") rowComInd5
    commentAFIS1Val    <- fmap toJSON $ lookupData id (T.pack "") rowComAfis1
    commentAFIS2Val    <- fmap toJSON $ lookupData id (T.pack "") rowComAfis2
    commentMOS1Val     <- fmap toJSON $ lookupData id (T.pack "") rowComMos1
    commentMOS2Val     <- fmap toJSON $ lookupData id (T.pack "") rowComMos2

 -- KPistruct
    let kpiValues = [ nbSitesVal, nbChannelsVal, nbCallsVal, nbMinutesVal, asrVal
                    , nerVal, postGADVal, attpsVal, afisVal, mosVal, pddVal
                    , csrVal, rtdVal, availVal, unavailVal, commentIndisp1Val
                    , commentIndisp2Val, commentIndisp3Val, commentIndisp4Val
                    , commentIndisp5Val, commentAFIS1Val  , commentAFIS2Val
                    , commentMOS1Val, commentMOS2Val]
    
    return kpiValues


-- take 2 decimals
trunc :: Double -> Double
trunc double = (fromInteger $ round $ double * (10^2)) / (10.0 ^^2)
toDouble :: T.Text -> Double
--toDouble xs = case (reads.chgComma.endBy "," $ xs :: [(Double,String)] ) of
toDouble xs = case T.double xs of  
    Right (d,s) -> trunc d
    Left _ -> 0

toInt :: T.Text -> Int
toInt xs = case T.decimal xs of 
    Right (d,s) -> d
    Left _ -> 0


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

servToBS :: [Pair] -> BL.ByteString 
servToBS  = encode . object -- (19% alloc)
--servToBS  = encodePretty . object -- (29% alloc)

data EnumVARIANT a      = EnumVARIANT
type IEnumVARIANT a     = IUnknown (EnumVARIANT a)
iidIEnumVARIANT :: IID (IEnumVARIANT ())
iidIEnumVARIANT = mkIID "{00020404-0000-0000-C000-000000000046}"

newEnum :: IDispatch a -> IO (Int, IEnumVARIANT b)
newEnum ip = do
  iunk  <- ip # propertyGet "_NewEnum" [] outIUnknown
  ienum <- iunk # queryInterface iidIEnumVARIANT
  len   <- ip   # propertyGet "Count" [] outInt
  return (len, castIface ienum)
  

getByOne ie def = do
    mb <- catchComException (ie # enumNextOne (fromIntegral sizeofVARIANT) resVariant) 
                        (\ _ -> return $ Just def)
    case mb of
	   Nothing -> return []
	   Just x -> do
	     xs <- getByOne ie def
	     return (x:xs)

enumVariants' :: Variant a => a -> IDispatch b -> IO (Int, [a])
enumVariants' def ip  = do
     (len, ienum) <- newEnum ip
--     enumNext (fromIntegral sizeofVARIANT) resVariant (fromIntegral len) ienum
     ls <- getByOne ienum def
     return (len, ls)
