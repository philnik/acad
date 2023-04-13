(push "c:/users/me/Application Data//acad/kernel/" asdf:*central-registry*)
;(push "z://lisp//acad/kernel/" asdf:*central-registry*)
(ql:quickload :acad)

(defun main ()
   (acad::hello)
  )
  


;(sb-ext:save-lisp-and-die #P"./test.exe" :toplevel #'main :compression nil :executable t)

