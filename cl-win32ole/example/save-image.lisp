
(in-package :cl-user)
(push "c:/users/me/quicklisp/local-projects/cl-win32ole/" asdf:*central-registry*)

;; (eval-when (:compile-toplevel :load-toplevel :execute)
;;   (asdf:oos 'asdf:load-op :cl-win32ole)
;;   (use-package :cl-win32ole))


(ql:quickload :cl-win32ole)
(ql:quickload :slynk)
(use-package :cl-win32ole)


(defun main ()
;  (sb-thread:make-thread
  (progn
    (ole-sys::co-initialize-apartment-threaded)
    ;;(co-initialize-apartment-threaded)
;     (co-initialize-multithreaded)
  (defparameter ie (create-object "InternetExplorer.Application"))
  (setf (ole ie :Visible) t)
  )
  )





(defun sbcl-save-sly-and-die ()
  "Save a sbcl image, even when running from inside Sly.
This function should only be used in the *inferior-buffer* buffer,
inside emacs."
  `(progn
  (mapcar #'(lambda (x)
              (slynk::close-connection 
               x nil nil))
          slynk::*connections*)
  (dolist (thread (remove
                   (slynk-backend::current-thread)
                   (slynk-backend::all-threads)))
    (slynk-backend::kill-thread thread))
  (sleep 1)
  ,(sb-ext:save-lisp-and-die #P"./your.exe"
                            :toplevel #'main
                            :executable t
                            :compression nil)
  ))

;; in *sly-inferior-lisp* buffer
(eval (sbcl-save-sly-and-die))
