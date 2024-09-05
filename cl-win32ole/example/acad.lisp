
(in-package :cl-user)


;; (eval-when (:compile-toplevel :load-toplevel :execute)
;;   (asdf:oos 'asdf:load-op :cl-win32ole)
;;   (use-package :cl-win32ole))

(defparameter cl_directory "c:/Users/filip/AppData/Roaming/lisp/cl-win32-sly")
(push cl_directory ql:*local-project-directories*)

(ql:quickload :cl-win32ole)
(use-package :cl-win32ole)
(use-package :cffi)
(in-package :cl-win32ole)


(defun ex ()
  (progn
    (cl-win32ole-sys::co-initialize)
    (defparameter acad (create-object "BricscadApp.AcadApplication"))
    (defparameter doc (ole acad :ActiveDocument))
    (defparameter model (ole doc :ModelSpace))
    (setf (ole acad :Visible) 1)
    (format t "acad.ActiveDocument.Name:~a~%" (ole acad :ActiveDocument :Name))

    (invoke model :AddLine "0,0" "100,100")
    (invoke model :ZoomAll)
    (cl-win32ole-sys::co-uninitialize)
    (list acad model doc)))

(setq mydata (ex))



(format t "acad.ActiveDocument.Name:~a~%" (ole acad :ActiveDocument :Name))
(invoke (ole doc :Application) :ZoomAll)
(defparameter application  (ole doc :Application))
(defparameter selectionset  (ole doc :SelectionSets))
(ole selectionset :Add "myselectionset")
(ole (ole doc :ModelSpace) :Count)



(setf d2 (make-array 3 :element-type 'double-float :initial-contents #(3.0d0 4.0d0 0.0d0)))
(setf v0 (alloc-variant))
(setf (variant-type v0) (logior VT_ARRAY VT_SAFEARRAY))
;(setf (variant-type v0) 4096)
(setf p1 (cffi:foreign-alloc :double :initial-contents d2))
(setf (variant-value v0) p1)
(to-lisp v0)
(variant-array-to-lisp v0)


(setf v1 (alloc-variant))
(setf (variant-type v1) VT_R8)
(setf (variant-value v1) 10.0d0)
(to-lisp v1)

(setf aaa(make-variant p1 8204))

(cffi:mem-aref p1 :double 0)
(cffi:mem-aref p1 :double 1)
(cffi:mem-aref p1 :double 2)

(cffi:mem-aref (variant-value v0) :double 0)
(cffi:mem-aref (variant-value v0) :double 1)
(cffi:mem-aref (variant-value v0) :double 2)
(cffi:mem-ref (variant-value v0) :double 0)

(invoke (ole doc :ModelSpace) :AddArc aaa 25.0 0.0 1.5)


(defparameter l '(1 2))

(loop for i from 1 to 20
      do (progn 
	   (setf ent1 (ole (ole doc :ModelSpace) :Item 2))
	   (format t "~a~%" (property ent1 :Layer))
	   ))
(elt selectionset 0)

      

(ole application :RunCommand "circle")
(ole application :RunCommand "0,0,0")
(ole application :RunCommand "(setq c (+ 1 2)")

(loop for i from 11 to 21
      do (ole application :RunCommand
	      (format nil "circle~%0,0,0~%~s~%" i)
	      )
      )




(ql:quickload "cffi")

(cffi:define-foreign-library winapi
  (:windows (:or "ole32.dll")))

(cffi:use-foreign-library winapi)

(cffi:defcfun ("CoInitialize" coinitialize) :void
  (arg :pointer))

(cffi:defcfun ("CoUninitialize" couninitialize) :void)

(couninitialize)

(cffi:defcfun ("CoInitialize" coinitialize) :int
  ((arg :pointer)))


(defun initialize-com ()
  (let ((result (coinitialize :void)))
    (if (zerop result)
        (format t "COM initialized successfully.~%")
        (format t "Failed to initialize COM. Error code: ~A~%" result))))

(initialize-com)

(defun uninitialize-com ()
  (couninitialize)
  (format t "COM uninitialized successfully.~%"))


(defun example-com-initialization ()
  (initialize-com)
  (unwind-protect
       (progn
         ;; Place your COM interactions here
         (format t "Performing COM operations...~%"))
    (uninitialize-com)))

;; Run the example function
(example-com-initialization)
