import System.Win32.Com 
import System.Win32.Com.Automation
--
-- createObjectExcel 
-- coming from Automation.hs and com.hs
--
xlWorkbookDefault, xlSaveChanges, xlDoNotSaveChanges :: Int
xlDoNotSaveChanges = 2
xlSaveChanges = 1
xlWorkbookDefault = 51

getActiveWB :: IDispatch a -> IO (IDispatch ()) 
getActiveWB obj = obj # propertyGet_0 "ActiveWorkBook"
--getRange  :: Variant b => String -> IDispatch a -> IO b 
getText :: IDispatch a -> IO String
getText obj= obj # propertyGet_0 "Text"
getRange  ::String -> IDispatch a -> IO (IDispatch ()) 
getRange rng obj = obj # propertyGet_1 "Range" rng


iidIDispatch_unsafe  = mkIID "{00020400-0000-0000-C000-000000000046}"

createObjExl :: IO (IDispatch ()) 
createObjExl = do
    clsidExcel <- clsidFromProgID "Excel.Application"
    pExl <- coCreateInstance clsidExcel  Nothing LocalProcess iidIDispatch_unsafe
    return pExl


fichierTest2 = "E:/Programmation/haskell/Com/qos1.xls"

main = coRun $ do 
    pExl <- createObjExl 
    workBooks <- pExl #  propertyGet_0 "Workbooks" 
    workBook <- workBooks #  propertyGet_1 "Open" fichierTest2
    workSheets <- workBook #  propertyGet_0 "Worksheets"
    sheetSel <- workSheets # propertyGet_1 "Item" (1::Int) :: IO (IDispatch ())
    text <- sheetSel # getRange "C1" ## getText 
    putStrLn text


    activeWBook <- pExl # getActiveWB
    activeWBook # method_1_0 "Save" xlWorkbookDefault
    workBooks # method_1_0 "Close"  xlSaveChanges
    pExl # method_0_0 "Quit"

    mapM release [sheetSel, workSheets,activeWBook,workBook, workBooks, pExl]
    --mapM release [workBook,workBooks, pExl]
    --mapM release [workBooks, pExl]
