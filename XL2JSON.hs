
{-# LANGUAGE OverloadedStrings , RecordWildCards #-}

import ExcelCom 
import RawToJSON
import KpiStructure
import qualified Data.Text as T 
import Data.Aeson.Types (Pair)
import Data.Aeson
import qualified Data.ByteString.Lazy.Char8 as BL

    
fichierTest = "C:/Users/lyup7323/Developpement/Haskell/Com/qos1.xls"
fichierTest2 = "E:/Programmation/haskell/Com/qos1.xls"
fichierTest3 = "E:/Programmation/haskell/Com/qos.xls"
fichierTest4 = "C:/Users/lyup7323/Developpement/Haskell/Com/qos.xls"


instance ToJSON KpiStruct where
    toJSON KpiStruct{..} = object [kpiName .= kpiData]

instance ToJSON ServStruct where
    toJSON ServStruct{..} = object [servName .= kpis]



sheetsName = ["BIV","BIC"]

main1 = coRun $ do
    (pExl, workBooks, workSheets) <- xlInit
    mapM_ (processRowData' workSheets) sheetsName
    xlQuit workBooks pExl


servToBS :: [Pair] -> BL.ByteString 
servToBS  = encode . object 
servToBS' :: Services -> BL.ByteString 
servToBS'  = encode . map toJSON

main = coRun $ do 
    (pExl, workBooks, workSheets) <- xlInit
    xs <- mapM (processRowData'' workSheets) sheetsName
    BL.writeFile "json.txt" $ servToBS' xs     
    

    xlQuit workBooks pExl

    
    
xlQuit workBooks appXl = do
    workBooks # method_1_0 "Close" xlDoNotSaveChanges
    appXl # method_0_0 "Quit"

xlInit = do   
    pExl <- createObjExl
    workBooks <- pExl # getWorkbooks
    pExl # propertySet "DisplayAlerts" [inBool False]
    workBook <- workBooks # openWorkBooks fichierTest3
    putStrLn  $"File loaded: " ++ fichierTest3
    workSheets <- workBook # getWSheets'
    return (pExl, workBooks, workSheets)
    

processRowData :: Sheet a -> String -> IO Pair
processRowData sheets sheetName = do 
    rowsService <- rowsFromSheet sheets sheetName
    return $ servToPair rowsService (T.pack sheetName) 
    
processRowData'' :: Sheet a -> String -> IO ServStruct
processRowData'' sheets sheetName= do 
    rowsService <- rowsFromSheet sheets sheetName
    return $ ServStruct (T.pack sheetName) (rawToStruct rowsService)
    
processRowData' :: Sheet a -> String -> IO ()
processRowData' sheets sheetName = do 
    rowsService <- rowsFromSheet sheets sheetName
    putStrLn $ "got all datas from " ++ sheetName
    printListData rowsService
    putStrLn $ replicate 50 '-'


rowsFromSheet :: Sheet a -> String -> IO [String] 
rowsFromSheet workSheets sheetName= do 
    sheetSel <- workSheets # propertyGet_1 "Item" sheetName

    let row = 79
    let lastrow =  "C7:BC"++ show row
    putStrLn $ "endrow = " ++  lastrow
    rng <- sheetSel # propertyGet_1 "Range" lastrow
    fmap snd $ rng # enumVariants


