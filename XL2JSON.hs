
{-# LANGUAGE OverloadedStrings , RecordWildCards #-}

import ExcelCom 
import RawToJSON
import qualified Data.Map as M (fromList, lookup)
import qualified Data.Text as T 
import Data.Aeson.Types (Pair,Value)
import Data.Aeson
import qualified Data.ByteString.Lazy.Char8 as BL

fichierTest = fichierTest4    
fichierTest1 = "C:/Users/lyup7323/Developpement/Haskell/Com/qos1.xls"
fichierTest2 = "E:/Programmation/haskell/Com/qos1.xls"
fichierTest3 = "E:/Programmation/haskell/Com/qos.xls"
fichierTest4 = "C:/Users/lyup7323/Developpement/Haskell/Com/qos.xls"



sheetsName = ["BIV","BIC"]


servToBS :: [Pair] -> BL.ByteString 
servToBS  = encode . object 


main = coRun $ do 
    (pExl, workBooks, workSheets) <- xlInit
    xs <- mapM (processRowData' workSheets) sheetsName
    BL.writeFile "json.txt" $ servToBS xs     

    xlQuit workBooks pExl

    
    
xlQuit workBooks appXl = do
    workBooks # method_1_0 "Close" xlDoNotSaveChanges
    appXl # method_0_0 "Quit"

xlInit = do   
    pExl <- createObjExl
    workBooks <- pExl # getWorkbooks
    pExl # propertySet "DisplayAlerts" [inBool False]
    workBook <- workBooks # openWorkBooks fichierTest
    putStrLn  $"File loaded: " ++ fichierTest
    workSheets <- workBook # getWSheets'
    return (pExl, workBooks, workSheets)
    

processRowData :: Sheet a -> String -> IO Pair
processRowData sheets sheetName = do 
    putStrLn $ "got all datas from " ++ sheetName
    putStrLn $ replicate 50 '-'
    rowsService <- rowsFromSheet sheets sheetName
    return $ servToPair rowsService (T.pack sheetName)  

processRowData' :: Sheet a -> String -> IO Pair
processRowData' sheets sheetName = do 
    putStrLn $ "got all datas from " ++ sheetName
    putStrLn $ replicate 50 '-'
    kpisVal <- valuesFromSheet sheets sheetName
    return $ servToPair' kpisVal (T.pack sheetName)    
    
valuesFromSheet :: Sheet a -> String -> IO [Value] 
valuesFromSheet workSheets sheetName= do 
    sheetSel <- workSheets # propertyGet_1 "Item" sheetName
    kpiNames <- mapM (\x -> sheetSel # getCells x 3 ## getText) [7..100]


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
        lookupData cast fill rowKpi  = maybe (return $ replicate 52 fill) -- default value [fill,...fill]
                                       (getData cast) -- handler 
                                       rowKpi -- Nothing or Just (kpi's row) 
        getData cast row = do 
            vals <- mapM (\x -> sheetSel # getCells row x ## getValue) [4..55]
            return $ fmap cast vals 

    nbSitesVal         <- fmap toJSON $ lookupData toInt 0 rowSite
    nbChannelsVal      <- fmap toJSON $ lookupData toInt 0 rowChannels
    nbMinutesVal       <- fmap toJSON $ lookupData toDouble 0.0 rowMinutes
    nbCallsVal         <- fmap toJSON $ lookupData toDouble 0.0 rowCalls
    postGADVal         <- fmap toJSON $ lookupData toDouble 0.0 rowPgad
    asrVal             <- fmap toJSON $  lookupData toDouble 0.0 rowAsr
    nerVal             <- fmap toJSON $ lookupData toDouble 0.0 rowNer
    attpsVal           <- fmap toJSON $ lookupData toDouble 0.0 rowAttps
    afisVal            <- fmap toJSON $ lookupData toDouble 0.0 rowAfis 
    mosVal             <- fmap toJSON $ lookupData toDouble 0.0 rowMos 
    pddVal             <- fmap toJSON $ lookupData toDouble 0.0 rowPdd
    csrVal             <- fmap toJSON $ lookupData toDouble 0.0 rowCsr 
    rtdVal             <- fmap toJSON $ lookupData toDouble 0.0 rowRtd 
    availVal           <- fmap toJSON $ lookupData toDouble 0.0 rowAvail 
    unavailVal         <- fmap toJSON $ lookupData toDouble 0.0 rowUnAvail 
    commentIndisp1Val  <- fmap toJSON $ lookupData id (""::String) rowComInd1 
    commentIndisp2Val  <- fmap toJSON $ lookupData id (""::String) rowComInd2
    commentIndisp3Val  <- fmap toJSON $ lookupData id (""::String) rowComInd3
    commentIndisp4Val  <- fmap toJSON $ lookupData id (""::String) rowComInd4
    commentIndisp5Val  <- fmap toJSON $ lookupData id (""::String) rowComInd5
    commentAFIS1Val    <- fmap toJSON $ lookupData id (""::String) rowComAfis1
    commentAFIS2Val    <- fmap toJSON $ lookupData id (""::String) rowComAfis2
    commentMOS1Val     <- fmap toJSON $ lookupData id (""::String) rowComMos1
    commentMOS2Val     <- fmap toJSON $ lookupData id (""::String) rowComMos2

 -- KPistruct
    let kpiValues = [ nbSitesVal, nbChannelsVal, nbCallsVal, nbMinutesVal, asrVal
                    , nerVal, postGADVal, attpsVal, afisVal, mosVal, pddVal
                    , csrVal, rtdVal, availVal, unavailVal, commentIndisp1Val
                    , commentIndisp2Val, commentIndisp3Val, commentIndisp4Val
                    , commentIndisp5Val, commentAFIS1Val  , commentAFIS2Val
                    , commentMOS1Val, commentMOS2Val]
    
    return kpiValues


    

rowsFromSheet :: Sheet a -> String -> IO [String] 
rowsFromSheet workSheets sheetName= do 
    sheetSel <- workSheets # propertyGet_1 "Item" sheetName

    let row = 79
    let lastrow =  "C7:BC"++ show row
    putStrLn $ "endrow = " ++  lastrow
    rng <- sheetSel # propertyGet_1 "Range" lastrow
    fmap snd $ rng # enumVariants


