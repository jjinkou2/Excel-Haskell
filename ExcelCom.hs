module ExcelCom (
        module System.Win32.Com
        , module System.Win32.Com.Automation
        , Range, Sheet
        , xlWorkbookDefault, xlSaveChanges, xlDoNotSaveChanges
        , xlUp, xlDown
        , getActiveCell
        , getText,getFormula, getValue, getInt, getDouble, getStr, getRanges, getCells0
        , setValue
        , select
        , showXl, hideXl
        , getActiveWB, getWorkbooks, openWorkBooks
        , getWSheets, getSheet, getSheetName, sheetSelect, getActiveSheet, getActiveSheet0
        , setSheetName
        , getRange, getCells
        ,createObjExl
        
        )


where


import System.Win32.Com 
import System.Win32.Com.Automation
import System.Win32.Com.HDirect.Pointer hiding ( freeBSTR )
import qualified System.Win32.Com.HDirect.Pointer as P ( freeBSTR )
import System.Win32.Com.HDirect.WideString

--
data Range_ a = Range__ 
type Range a = IDispatch (Range_ a)

data Sheet_ a = Sheet__ 
type Sheet a = IDispatch (Sheet_ a)

xlWorkbookDefault, xlSaveChanges, xlDoNotSaveChanges :: Int
xlDoNotSaveChanges = 2
xlSaveChanges = 1
xlWorkbookDefault = 51

xlUp, xlDown :: Int
xlDown = 0xffffefe7
xlToLeft = 0xffffefc1
xlToRight = 0xffffefbf
xlUp = 0xffffefbe

--getActiveCell :: IDispatch a -> IO (Range ()) 
getActiveCell :: IDispatch a -> IO (IDispatch ()) 
getActiveCell obj = obj # propertyGet_0 "ActiveCell"


getFormula :: IDispatch a -> IO String
getFormula obj= obj # propertyGet_0 "Formula"
getText :: IDispatch a -> IO String
getText obj= obj # propertyGet_0 "Text"

setValue :: String->IDispatch a -> IO ()
setValue str obj= obj # propertySet_1 "Value" str

getValue :: Variant b => IDispatch a -> IO b
getValue obj = obj # propertyGet_0 "Value"

getStr :: IDispatch a -> IO String
getStr = getValue 

getInt :: IDispatch a -> IO Int
getInt = getValue 

getDouble :: IDispatch a -> IO Double
getDouble = getValue  

getRanges :: (Variant a1, Variant a2) => a1 -> a2 -> IDispatch a -> IO (IDispatch ())
getRanges cell1 cell2 =
  propertyGet "Range" [ inVariant cell1 , inVariant cell2 ] outIUnknown

getCells0 :: IDispatch a0 -> IO (Range ())
getCells0 obj = obj # propertyGet "Cells" [] outIUnknown

--select :: (Variant a1)=>IDispatch () -> IO  (a1)
--select = function1 "Select" [] outVariant

select :: (Variant a0)=>IDispatch a0 -> IO () 
select obj = obj # function1 "Select" [] outVariant

{-
getSelection :: IDispatch a -> IO (IDispatch ()) 
getSelection obj = obj # propertyGet_0 "Selection"


setText :: String->IDispatch a -> IO ()
setText str obj= obj # propertySet_1 "Text" str
-}

showXl obj = obj # propertySet "Visible" [inBool True]
hideXl obj = obj # propertySet "Visible" [inBool False]

-- WorkBooks
getActiveWB :: IDispatch a -> IO (IDispatch ()) 
getActiveWB obj = obj # propertyGet_0 "ActiveWorkBook"
getWorkbooks :: IDispatch a -> IO (IDispatch ()) 
getWorkbooks obj = obj # propertyGet_0 "Workbooks"

openWorkBooks :: String -> IDispatch a -> IO (IDispatch())
openWorkBooks fp obj = obj # propertyGet_1 "Open" fp 

-- sheets
getWSheets :: IDispatch a -> IO (Sheet a) 
getWSheets obj = obj # propertyGet_0 "Worksheets"

getSheet :: (Variant a0,Variant a1)=> a0  -> IDispatch a1 -> IO (IDispatch ()) 
getSheet n obj = obj # getWSheets ## propertyGet_1 "Item" n

sheetSelect :: (Variant a0,Variant a1) => a0  -> IDispatch a1 -> IO ()
sheetSelect n obj = obj # getSheet n ## select

--getActiveSheet0 :: IDispatch a0 -> IO (IDispatch ())
getActiveSheet0 :: IDispatch a0 -> IO (IDispatch (Sheet a))
getActiveSheet0 = propertyGet "ActiveSheet" [] outIDispatch

getActiveSheet :: IDispatch a -> IO (IDispatch (Sheet a))  
getActiveSheet obj = obj  # propertyGet_0 "ActiveSheet" 

setSheetName :: IDispatch (Sheet a) -> String ->  IO ()
setSheetName obj str = obj # propertySet_1 "Name" str

getSheetName :: IDispatch (Sheet a) -> IO String
getSheetName obj = obj # propertyGet_0 "Name"

-- Cells
--getRange  :: Variant b => String -> IDispatch a -> IO b 
getRange  ::String -> IDispatch a -> IO (IDispatch ()) 
getRange rng obj = obj # propertyGet_1 "Range" rng

getCells :: Int -> Int -> IDispatch a -> IO (Range ())
getCells col row obj = obj # propertyGet_2 "Cells" col row

--
-- createObjectExcel 
-- coming from Automation.hs and com.hs
--
iidAppl  = mkIID "{00020400-0000-0000-C000-000000000046}"
iidIDispatch_unsafe  = mkIID "{00020400-0000-0000-C000-000000000046}"

createObjExl :: IO (IDispatch ()) 
createObjExl = do
    clsidExcel <- clsidFromProgID "Excel.Application"
    pExl <- coCreateInstance clsidExcel  Nothing LocalProcess iidAppl
    return pExl



getFileXl fname  = do
    pf <- coCreateObject' "Excel.Application" iidIPersistFile
    stackWideString fname $ \pfname -> do
      persistfileLoad' pf pfname 0
      pf # queryInterface iidAppl

coCreateObject' :: ProgID -> IID (IUnknown a) -> IO (IUnknown a)
coCreateObject' progid iid = do
    clsid  <- clsidFromProgID progid
    coCreateInstance clsid Nothing LocalProcess iid




