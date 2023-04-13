(push "c:/users/me/Application Data/lisp/cl-win32ole/" asdf:*central-registry*)
(eval-when (:compile-toplevel :load-toplevel :execute)
  (asdf:oos 'asdf:load-op :cl-win32ole)
  (use-package :cl-win32ole))

(defun wsh-example ()
  (let ((wsh (create-object "BricscadApp.AcadApplication")))
    (ole wsh :ActiveDocument)))

(wsh-example)
