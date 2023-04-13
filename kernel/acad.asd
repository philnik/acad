(push "c:/users/me/Application Data/acad/cl-win32ole/" asdf:*central-registry*)
;(push "z:/lisp/acad/cl-win32ole/" asdf:*central-registry*)
(ql:quickload :cl-annot)
(defsystem "acad"
  :depends-on (:cl-win32ole :cl-annot)
  :components (
	       (:file "acad")
	       )
  )
