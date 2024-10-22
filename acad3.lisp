(in-package :cl-user)
;; ;; ;; (eval-when (:compile-toplevel :load-toplevel :execute)
;; ;; (asdf:oos 'asdf:load-op :cl-win32ole)
;; ;; (asdf:oos 'asdf:load-op :cl-win32ole-sys)
;; ;; ;;   (use-package :cl-win32ole))
					;
					; (defparameter *cl_directory* "c:/Users/filip/AppData/Roaming/lisp/acad/cl-win32ole/")
;; (push *cl_directory* asdf:*central-registry*)

;; ;(ql:quickload 'cl-win32ole)
;; ;(use-package 'cl-win32ole)
;; ;(use-package 'cl-win32ole-sys)
;; ;(in-package :cl-win32ole)
;; ;; ;(push cl_directory ql:*local-project-directories*)
;; ;; ;(ql:quickload :cl-win32ole)
;; (load (concatenate 'string *cl_directory* "cl-win32ole.asd"))
;; (load (concatenate 'string *cl_directory* "cl-win32ole-sys.asd"))
;; (load (concatenate 'string *cl_directory* "sys/dll.lisp"))
;; (load (concatenate 'string *cl_directory*  "sys/ffi.lisp"))

(format t "~a" (asdf:asdf-version))
(defparameter *cl_win32ole_directory* "c:/Users/filip/AppData/Roaming/lisp/acad/cl-win32ole/")
;(defparameter *cl_win32ole_directory* "c:/Users/f.nikolakopoulos/AppData/Roaming/lisp/acad/cl-win32ole/")
(push *cl_win32ole_directory* asdf:*central-registry*)
(asdf:load-system :cl-win32ole)
;;(asdf:load-system 'com.google.base)
(ql:quickload 'com.google.base)
(ql:quickload 'swank)
(ql:quickload 'slynk)
(defparameter *emacs-port* 4005)
(defparameter *swank-client-port* 10000)


(use-package :cl-win32ole)
(use-package :cffi)
(in-package :cl-win32ole)


(defun swank-thread ()
  "Returns a thread that's acting as a Swank server."
  (dolist (thread (sb-thread:list-all-threads))
    (when (com.google.base:prefixp "Swank" (sb-thread:thread-name thread))
      (return thread))))

(defun wait-for-swank-thread ()
  "Wait for the Swank server thread to exit."
  (let ((swank-thread (swank-thread)))
    (when swank-thread
      (sb-thread:join-thread swank-thread))))


(defun ex ()
  (progn
    (cl-win32ole-sys::coinitialize)
    (cl-win32ole-sys::co-initialize-multithreaded-ex)
    (defparameter acad (create-object "BricscadApp.AcadApplication" 0)))
    (defparameter doc (ole acad :ActiveDocument))
    (defparameter model (ole doc :ModelSpace))
    (setf (ole acad :Visible) 1)
    

    (setf swank:*configure-emacs-indentation* nil
        swank::*enable-event-ghistory* nil
        swank:*log-events* t)
    (defparameter *emacs-port* 4005)
    (defparameter *swank-client-port* 10000)
  
    (swank:create-server :port *emacs-port* :dont-close t)
    (swank:create-server :port *swank-client-port* :dont-close t)
    (wait-for-swank-thread)
  )
    
  ;;;; on emacs we run
;;(slime-connect "localhost" 4005)
;;
;;on new slime connection: (in-package :cl-win32ole)
;;to get access on local variables


(setq mydata (ex))
;;(setq mydata (ex 1))


