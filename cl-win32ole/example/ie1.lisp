
(in-package :cl-user)


;; (eval-when (:compile-toplevel :load-toplevel :execute)
;;   (asdf:oos 'asdf:load-op :cl-win32ole)
;;   (use-package :cl-win32ole))


(ql:quickload :cl-win32ole)
(use-package :cl-win32ole)


(defun ie-example2 ()
  (co-initialize)
  (let ((ie (create-object "InternetExplorer.Application")))
    (setf (ole ie :visible) t)
    (ole ie :navigate "http://www.google.co.jp/")
    (loop while (ole ie :busy) do (sleep 0.5))
    (setf (ole ie :document :all :item "q" :value) "Common Lisp")
    (ole ie :document :all :item "btnG" :click)
    (sleep 3)
    (ole ie :quit)))

(ie-example2)

(defun ie-example3 ()
  (co-initialize-multithreaded)
  (let ((ie (create-object "InternetExplorer.Application")))
    (with-slots (visible busy document) ie
      (setf visible t)
      (invoke ie :navigate "http://www.yahoo.co.jp/")
      (loop while busy do (sleep 0.5))
      (with-slots (all) document
        (print (invoke all :tags "A")))
      (ole ie :quit))))
