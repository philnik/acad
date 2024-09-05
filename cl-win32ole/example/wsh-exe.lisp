
;;; push cl-win32ole to asdf:*central-registry or add to quicklisp/localprojects
;(push ".../cl-win32ole/" asdf:*central-registry*)




(ql:quickload :cl-win32ole)
(use-package :cl-win32ole)


(defun wsh-example ()
  (co-initialize-multithreaded)
  (let ((wsh (create-object "WScript.Shell")))
    (invoke wsh :popup "Hello")))


(sb-ext:save-lisp-and-die #P"./wsh.exe"
			  :executable t
			  :toplevel #'wsh-example
			  :compression nil)
