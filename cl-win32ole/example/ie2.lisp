
(in-package :cl-user)


;; (eval-when (:compile-toplevel :load-toplevel :execute)
;;   (asdf:oos 'asdf:load-op :cl-win32ole)
;;   (use-package :cl-win32ole))


(ql:quickload :cl-win32ole)
(use-package :cl-win32ole)


(co-initialize-multithreaded)
(defparameter ie (create-object "InternetExplorer.Application"))
(setf (ole ie :Visible) t)
(ql:quickload :slynk)
(slynk:create-server :port 4008)



