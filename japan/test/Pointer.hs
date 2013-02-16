-----------------------------------------------------------------------------
-- |
-- Module      :  System.Win32.Com.HDirect.Pointer
-- Copyright   :  (c) Daan Leijen, leijen@@fwi.uva.nl 1998
-- License     :  BSD-style (see the file libraries/base/LICENSE)
-- 
-- Maintainer  :  sof@forkIO.com
-- Stability   :  provisional
-- Portability :  portable
--
-- This module is part of HaskellDirect (H\/Direct), providing 
-- helper functions over Ptrs + allocation\/freeing of memory via
-- malloc\/free or the COM task allocator.
--
-----------------------------------------------------------------------------
module Pointer  
	( 
	  Ptr
	
	, allocMemory
	, stackFrame

	, writeSeqAtDec

        , freeMemory
	, freeBSTR
	, freeWith
	, freeWithC
	
	, primNoFree
	
	, finalNoFree
	, finalFreeMemory
	
	, makeFO

       ) where

import Foreign.Ptr
import Foreign.ForeignPtr
import PointerPrim
import Data.Word     ( Word32 )
import Control.Monad
import Control.Exception (catch)


type Finalizer a  = Ptr a -> IO ()

makeFO :: Ptr a -> FunPtr (Ptr a -> IO ()) -> IO (ForeignPtr b)
makeFO obj finaliser = newForeignPtr (mkFinal finaliser obj) obj >>= return.castForeignPtr

mkFinal final _ = final

--Helpers.

writeSeqAtDec :: Word32 -> [Ptr a -> IO ()] -> Ptr a -> IO ()
writeSeqAtDec size ws ptr = go init_ptr ws
  where
   len           = fromIntegral (length ws - 1)
   init_ptr      = ptr `plusPtr` (size_i * len)
   size_i        = fromIntegral size

   go _   []     = return ()
   go ptr (x:xs) = do
      x ptr
      let ptr_next = ptr `plusPtr` (-size_i)
      go ptr_next xs


-- | Use 'stackFrame' when the allocated chunk have a 
-- limited and known lifetime.
stackFrame :: Word32 -> (Ptr a -> IO b) -> IO b
stackFrame size f
      = do p <- allocMemory size
           f p `always` primFreeMemory (castPtr p)


-- Special free routines for pointers. Use them to manually free pointers.

freeMemory            = freeWithC primFreeMemory
freeBSTR              = freeWithC primFreeBSTR

freeWithC :: Finalizer () -> Ptr a -> IO ()
freeWithC final p = final (castPtr p)

freeWith :: (Ptr a -> IO ()) -> Ptr a -> IO ()
freeWith free p = free p 

-- Helper functions that doesn't really have a good home to go to:

always :: IO a -> IO () -> IO a
always io action
      = do x <- io `catch` (\ err -> do { action; ioError err })
           action
           return x

--Primitives/helpers:

allocMemory :: Word32 -> IO (Ptr a)
allocMemory sz = do
  a <- primAllocMemory sz
  if a == nullPtr then
     ioError (userError "allocMemory: not enough memory")
   else
     return (castPtr a)
