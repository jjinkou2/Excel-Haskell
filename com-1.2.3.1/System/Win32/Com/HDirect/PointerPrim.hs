{-# OPTIONS_GHC -XCPP -fglasgow-exts -#include "PointerSrc.h" #-}
-- Automatically generated by HaskellDirect (ihc.exe), snapshot 171208
-- Created: 23:37 Pacific Standard Time, Wednesday 17 December, 2008
-- Command line: -fno-qualified-names -fkeep-hresult -fout-pointers-are-not-refs -c System/Win32/Com/HDirect/PointerPrim.idl -o System/Win32/Com/HDirect/PointerPrim.hs

module System.Win32.Com.HDirect.PointerPrim
       ( primNoFree
       , primFreeBSTR
       , primFreeMemory
       , finalNoFree
       , finalFreeMemory
       , primAllocMemory
       , primFinalise
       ) where
       
import Prelude
import Data.Word (Word32)
import Foreign.Ptr (Ptr)
import System.IO.Unsafe (unsafePerformIO)

primNoFree :: Ptr ()
           -> IO ()
primNoFree p =
  prim_System_Win32_Com_HDirect_PointerPrim_primNoFree p

foreign import ccall "primNoFree" prim_System_Win32_Com_HDirect_PointerPrim_primNoFree :: Ptr () -> IO ()
primFreeBSTR :: Ptr ()
             -> IO ()
primFreeBSTR p =
  prim_System_Win32_Com_HDirect_PointerPrim_primFreeBSTR p

foreign import ccall "primFreeBSTR" prim_System_Win32_Com_HDirect_PointerPrim_primFreeBSTR :: Ptr () -> IO ()
primFreeMemory :: Ptr ()
               -> IO ()
primFreeMemory p =
  prim_System_Win32_Com_HDirect_PointerPrim_primFreeMemory p

foreign import ccall "primFreeMemory" prim_System_Win32_Com_HDirect_PointerPrim_primFreeMemory :: Ptr () -> IO ()
finalNoFree :: Ptr ()
finalNoFree =
  unsafePerformIO (prim_System_Win32_Com_HDirect_PointerPrim_finalNoFree)

foreign import ccall "finalNoFree" prim_System_Win32_Com_HDirect_PointerPrim_finalNoFree :: IO (Ptr ())
finalFreeMemory :: Ptr ()
finalFreeMemory =
  unsafePerformIO (prim_System_Win32_Com_HDirect_PointerPrim_finalFreeMemory)

foreign import ccall "finalFreeMemory" prim_System_Win32_Com_HDirect_PointerPrim_finalFreeMemory :: IO (Ptr ())
primAllocMemory :: Word32
                -> IO (Ptr ())
primAllocMemory sz =
  prim_System_Win32_Com_HDirect_PointerPrim_primAllocMemory sz

foreign import ccall "primAllocMemory" prim_System_Win32_Com_HDirect_PointerPrim_primAllocMemory :: Word32 -> IO (Ptr ())
primFinalise :: Ptr ()
             -> Ptr ()
             -> IO ()
primFinalise finaliser finalisee =
  prim_System_Win32_Com_HDirect_PointerPrim_primFinalise finaliser
                                                         finalisee

foreign import ccall "primFinalise" prim_System_Win32_Com_HDirect_PointerPrim_primFinalise :: Ptr () -> Ptr () -> IO ()

