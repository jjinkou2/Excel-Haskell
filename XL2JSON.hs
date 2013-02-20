
import ExcelCom 
import RawToJSON

    
fichierTest = "C:/Users/lyup7323/Developpement/Haskell/Com/qos1.xls"
fichierTest2 = "E:/Programmation/haskell/Com/qos1.xls"
fichierTest3 = "E:/Programmation/haskell/Com/qos.xls"
fichierTest4 = "C:/Users/lyup7323/Developpement/Haskell/Com/qos.xls"




sheetsName = ["BIV","BIC"]

main = coRun $ do
    (pExl, workBooks, workSheets) <- xlInit
    mapM_ (processRowData workSheets) sheetsName
    xlQuit workBooks pExl

xlQuit workBooks appXl = do
    workBooks # method_1_0 "Close" xlDoNotSaveChanges
    appXl # method_0_0 "Quit"

xlInit = do   
    pExl <- createObjExl
    workBooks <- pExl # getWorkbooks
    pExl # propertySet "DisplayAlerts" [inBool False]
    workBook <- workBooks # openWorkBooks fichierTest4
    putStrLn  $"File loaded: " ++ fichierTest4
    workSheets <- workBook # getWSheets'
    return (pExl, workBooks, workSheets)
    

processRowData :: Sheet a -> String -> IO ()
processRowData sheets sheetName = do 
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


