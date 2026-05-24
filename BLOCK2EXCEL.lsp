;; =============================================================================
;; BLOCK2EXCEL.lsp - Извлечение координат блоков из чертежа AutoCAD в Excel
;; =============================================================================
;; При загрузке файла автоматически создаются определения блоков:
;;   B2E_POINT  - точка с автонумерацией (кружок + атрибут NUM)
;;   B2E_ORIGIN - маркер начальной точки (квадрат с диагоналями)
;;
;; Команды:
;;   B2E                  - открыть панель инструментов (главное меню)
;;   B2E_INSERT           - вставка блока B2E_POINT с автоматическим номером
;;   B2E_ORIGIN           - указать начальную точку (точку отсчёта)
;;   B2E_RESET            - сброс/установка значения счётчика нумерации
;;   B2E_SETTINGS         - открыть окно настроек
;;   B2E_INSTALL_PANEL    - установить панель инструментов AutoCAD
;;   B2E_UNINSTALL_PANEL  - удалить панель инструментов
;;   BLOCK2EXCEL          - выбор блоков рамкой/кликом и выгрузка в CSV
;;   BLOCK2EXCEL_ALL      - выгрузка ВСЕХ вхождений блоков на чертеже
;;   BLOCK2EXCEL_NAME     - выгрузка блоков по имени (с подстановкой *)
;;
;; Перед записью в CSV координаты пересчитываются по формуле:
;;   X_итог = X_нач - dX - X_точки
;;   Y_итог = Y_нач - dY - Y_точки
;; где X_нач/Y_нач - координаты начальной точки, dX/dY - поправки.
;;
;; Все параметры хранятся в .dwg в системных переменных:
;;   USERI5 - счётчик нумерации
;;   USERR1 - масштаб блоков
;;   USERR2 - X начальной точки
;;   USERR3 - Y начальной точки
;;   USERR4 - поправка dX
;;   USERR5 - поправка dY
;;
;; Кодировка CSV: UTF-8 с BOM, разделитель ";" (для русской локали Excel)
;; =============================================================================

(vl-load-com)

;; Имена блоков
(setq *b2e:block-name*  "B2E_POINT"
      *b2e:origin-name* "B2E_ORIGIN")

;; ===========================================================================
;; НАСТРОЙКИ (хранятся в системных переменных)
;; ===========================================================================

(defun b2e:init-defaults ( / )
  (if (zerop (getvar "USERR1")) (setvar "USERR1" 1.0))
  (princ)
)

(defun b2e:get-scale  () (getvar "USERR1"))
(defun b2e:set-scale  (v) (setvar "USERR1" (float v)))

(defun b2e:get-origin-x () (getvar "USERR2"))
(defun b2e:get-origin-y () (getvar "USERR3"))
(defun b2e:set-origin (x y)
  (setvar "USERR2" (float x))
  (setvar "USERR3" (float y))
)

(defun b2e:get-corr-x () (getvar "USERR4"))
(defun b2e:get-corr-y () (getvar "USERR5"))
(defun b2e:set-corr (dx dy)
  (setvar "USERR4" (float dx))
  (setvar "USERR5" (float dy))
)

;; ===========================================================================
;; УТИЛИТЫ
;; ===========================================================================

(defun b2e:csv-esc (val / s)
  (setq s (cond ((null val) "")
                ((numberp val) (rtos val 2 4))
                ((listp val) (strcat (rtos (car val) 2 4) ","
                                     (rtos (cadr val) 2 4)))
                (T (vl-princ-to-string val))))
  (if (or (vl-string-search ";" s)
          (vl-string-search "\"" s)
          (vl-string-search "\n" s))
    (strcat "\"" (b2e:replace-all s "\"" "\"\"") "\"")
    s)
)

(defun b2e:replace-all (str old new / pos len-old result)
  (setq result "" len-old (strlen old))
  (while (setq pos (vl-string-search old str))
    (setq result (strcat result (substr str 1 pos) new)
          str (substr str (+ pos len-old 1))))
  (strcat result str)
)

(defun b2e:get-attrs (ename / e attrs result tag val)
  (setq result "" attrs '())
  (setq e (entnext ename))
  (while (and e (/= (cdr (assoc 0 (entget e))) "SEQEND"))
    (if (= (cdr (assoc 0 (entget e))) "ATTRIB")
      (progn
        (setq tag (cdr (assoc 2 (entget e)))
              val (cdr (assoc 1 (entget e))))
        (setq attrs (cons (cons tag val) attrs))))
    (setq e (entnext e)))
  (setq attrs (reverse attrs))
  (foreach pair attrs
    (setq result
      (if (= result "")
        (strcat (car pair) "=" (cdr pair))
        (strcat result " | " (car pair) "=" (cdr pair)))))
  result
)

(defun b2e:get-attr-by-tag (ename tag / e found result)
  (setq result "" found nil)
  (setq e (entnext ename))
  (while (and e (not found)
              (/= (cdr (assoc 0 (entget e))) "SEQEND"))
    (if (and (= (cdr (assoc 0 (entget e))) "ATTRIB")
             (= (strcase (cdr (assoc 2 (entget e)))) (strcase tag)))
      (progn
        (setq result (cdr (assoc 1 (entget e))))
        (setq found T)))
    (setq e (entnext e)))
  result
)

;; ===========================================================================
;; ИЗВЛЕЧЕНИЕ ДАННЫХ БЛОКА (с пересчётом координат)
;; ===========================================================================
;; Формула: X_итог = X_нач - dX - X_точки
;;          Y_итог = Y_нач - dY - Y_точки

(defun b2e:get-block-data (ename / ent name pt layer rot scale attrs
                                    num x-orig y-orig dx dy x-new y-new)
  (setq ent (entget ename))
  (if (= (cdr (assoc 0 ent)) "INSERT")
    (progn
      (setq name (vlax-get-property
                   (vlax-ename->vla-object ename) 'EffectiveName))
      (setq pt    (cdr (assoc 10 ent))
            layer (cdr (assoc 8  ent))
            rot   (cdr (assoc 50 ent))
            scale (cdr (assoc 41 ent))
            attrs (b2e:get-attrs ename)
            num   (b2e:get-attr-by-tag ename "NUM"))
      (setq x-orig (b2e:get-origin-x)
            y-orig (b2e:get-origin-y)
            dx     (b2e:get-corr-x)
            dy     (b2e:get-corr-y))
      (setq x-new (- x-orig dx (car pt))
            y-new (- y-orig dy (cadr pt)))
      (list name
            num                              ; № (атрибут NUM)
            x-new                            ; X пересчитанный
            y-new                            ; Y пересчитанный
            (car pt)                         ; X исходный
            (cadr pt)                        ; Y исходный
            layer
            (* (/ rot pi) 180.0)             ; поворот, °
            scale
            attrs))
    nil)
)

(defun b2e:write-csv (filepath rows / fp row cell first)
  (setq fp (open filepath "w"))
  (if (null fp)
    (progn (princ (strcat "\nОшибка: не могу создать файл " filepath)) nil)
    (progn
      (write-char 239 fp) (write-char 187 fp) (write-char 191 fp)
      (write-line
        (strcat "# Начальная точка: X="
                (rtos (b2e:get-origin-x) 2 4)
                " Y=" (rtos (b2e:get-origin-y) 2 4)
                "; Поправка: dX=" (rtos (b2e:get-corr-x) 2 4)
                " dY=" (rtos (b2e:get-corr-y) 2 4)
                "; Формула: X_итог = X_нач - dX - X_точки")
        fp)
      (write-line
        (strcat "Имя блока;№;X (расчёт);Y (расчёт);"
                "X (исх.);Y (исх.);Слой;Поворот (град);Масштаб;Атрибуты")
        fp)
      (foreach row rows
        (setq first T)
        (foreach cell row
          (if first
            (progn (princ (b2e:csv-esc cell) fp) (setq first nil))
            (progn (princ ";" fp) (princ (b2e:csv-esc cell) fp))))
        (write-line "" fp))
      (close fp)
      T))
)

(defun b2e:get-save-path (/ path)
  (setq path (getfiled "Сохранить координаты блоков как..."
                       (strcat (getvar "DWGPREFIX") "blocks_coords.csv")
                       "csv" 1))
  path
)

(defun b2e:open-in-excel (filepath)
  (if filepath
    (progn
      (vl-cmdf "_.SHELL" (strcat "start \"\" \"" filepath "\""))
      (princ (strcat "\nФайл сохранён: " filepath))))
)

;; ===========================================================================
;; ОПРЕДЕЛЕНИЯ БЛОКОВ
;; ===========================================================================

(defun b2e:block-exists (name)
  (not (null (tblsearch "BLOCK" name)))
)

(defun b2e:ensure-point-block ( / doc blocks blk)
  (setq doc    (vla-get-ActiveDocument (vlax-get-acad-object))
        blocks (vla-get-Blocks doc))
  (if (not (b2e:block-exists *b2e:block-name*))
    (progn
      (princ (strcat "\nСоздаю блок " *b2e:block-name* "..."))
      (setq blk (vla-Add blocks
                  (vlax-3d-point '(0.0 0.0 0.0))
                  *b2e:block-name*))
      (vla-AddCircle blk (vlax-3d-point '(0.0 0.0 0.0)) 2.5)
      (vla-AddLine blk
        (vlax-3d-point '(-2.5 0.0 0.0))
        (vlax-3d-point '( 2.5 0.0 0.0)))
      (vla-AddLine blk
        (vlax-3d-point '(0.0 -2.5 0.0))
        (vlax-3d-point '(0.0  2.5 0.0)))
      (vla-AddAttribute blk
        2.5
        acAttributeModeVerify
        "Номер точки:"
        (vlax-3d-point '(3.5 3.5 0.0))
        "NUM"
        "1")
      T)
    T)
)

(defun b2e:ensure-origin-block ( / doc blocks blk)
  (setq doc    (vla-get-ActiveDocument (vlax-get-acad-object))
        blocks (vla-get-Blocks doc))
  (if (not (b2e:block-exists *b2e:origin-name*))
    (progn
      (princ (strcat "\nСоздаю блок " *b2e:origin-name* "..."))
      (setq blk (vla-Add blocks
                  (vlax-3d-point '(0.0 0.0 0.0))
                  *b2e:origin-name*))
      ;; Квадрат 5x5
      (vla-AddLine blk
        (vlax-3d-point '(-2.5 -2.5 0.0))
        (vlax-3d-point '( 2.5 -2.5 0.0)))
      (vla-AddLine blk
        (vlax-3d-point '( 2.5 -2.5 0.0))
        (vlax-3d-point '( 2.5  2.5 0.0)))
      (vla-AddLine blk
        (vlax-3d-point '( 2.5  2.5 0.0))
        (vlax-3d-point '(-2.5  2.5 0.0)))
      (vla-AddLine blk
        (vlax-3d-point '(-2.5  2.5 0.0))
        (vlax-3d-point '(-2.5 -2.5 0.0)))
      ;; Диагонали
      (vla-AddLine blk
        (vlax-3d-point '(-2.5 -2.5 0.0))
        (vlax-3d-point '( 2.5  2.5 0.0)))
      (vla-AddLine blk
        (vlax-3d-point '(-2.5  2.5 0.0))
        (vlax-3d-point '( 2.5 -2.5 0.0)))
      (vla-AddAttribute blk
        2.0
        acAttributeModeInvisible
        "Метка:"
        (vlax-3d-point '(3.5 3.5 0.0))
        "LABEL"
        "ORIGIN")
      T)
    T)
)

;; ===========================================================================
;; СЧЁТЧИК НУМЕРАЦИИ
;; ===========================================================================

(defun b2e:counter-get () (getvar "USERI5"))
(defun b2e:counter-set (n) (setvar "USERI5" n) n)
(defun b2e:counter-next ( / n)
  (setq n (1+ (b2e:counter-get)))
  (b2e:counter-set n)
  n
)

(defun b2e:set-attr-value (block-ref tag value / atts att)
  (setq atts (vlax-invoke block-ref 'GetAttributes))
  (foreach att atts
    (if (= (strcase (vla-get-TagString att)) (strcase tag))
      (vla-put-TextString att value)))
)

;; ===========================================================================
;; КОМАНДЫ
;; ===========================================================================

;; --- Вставка точки с автонумерацией -------------------------------------

(defun c:B2E_INSERT ( / doc ms pt num block-ref continue scale)
  (b2e:ensure-point-block)
  (setq doc      (vla-get-ActiveDocument (vlax-get-acad-object))
        ms       (vla-get-ModelSpace doc)
        scale    (b2e:get-scale)
        continue T)
  (princ (strcat "\nТекущий счётчик: " (itoa (b2e:counter-get))
                 "  Масштаб блоков: " (rtos scale 2 3)))
  (princ "\nУкажите точки вставки (Esc - выход):")
  (while continue
    (setq pt (getpoint "\nТочка вставки: "))
    (if (null pt)
      (setq continue nil)
      (progn
        (setq num (b2e:counter-next))
        (setq block-ref
          (vla-InsertBlock ms
            (vlax-3d-point pt)
            *b2e:block-name*
            scale scale scale
            0.0))
        (b2e:set-attr-value block-ref "NUM" (itoa num))
        (princ (strcat "\nВставлен блок №" (itoa num)
                       " в точку ("
                       (rtos (car pt) 2 3) ", "
                       (rtos (cadr pt) 2 3) ")")))))
  (princ)
)

;; --- Установка начальной точки -----------------------------------------

(defun c:B2E_ORIGIN ( / doc ms pt scale ss)
  (b2e:ensure-origin-block)
  (setq doc   (vla-get-ActiveDocument (vlax-get-acad-object))
        ms    (vla-get-ModelSpace doc)
        scale (b2e:get-scale))
  (princ (strcat "\nТекущая начальная точка: X="
                 (rtos (b2e:get-origin-x) 2 4)
                 " Y=" (rtos (b2e:get-origin-y) 2 4)))
  (setq pt (getpoint "\nУкажите начальную точку (точку отсчёта): "))
  (if pt
    (progn
      ;; Удалим старые маркеры B2E_ORIGIN
      (setq ss (ssget "_X" (list '(0 . "INSERT")
                                  (cons 2 *b2e:origin-name*))))
      (if ss
        (progn
          (princ "\nУдаляю предыдущий маркер начальной точки...")
          (repeat (sslength ss)
            (entdel (ssname ss 0))
            (if (> (sslength ss) 1) (ssdel (ssname ss 0) ss)))))
      ;; Сохраняем координаты
      (b2e:set-origin (car pt) (cadr pt))
      ;; Вставляем маркер
      (vla-InsertBlock ms
        (vlax-3d-point pt)
        *b2e:origin-name*
        scale scale scale
        0.0)
      (princ (strcat "\nНачальная точка установлена: X="
                     (rtos (car pt) 2 4)
                     " Y=" (rtos (cadr pt) 2 4))))
    (princ "\nОтменено."))
  (princ)
)

;; --- Сброс счётчика ------------------------------------------------------

(defun c:B2E_RESET ( / start)
  (princ (strcat "\nТекущее значение счётчика: "
                 (itoa (b2e:counter-get))))
  (setq start (getint "\nС какого номера начинать [0]: "))
  (if (null start) (setq start 0))
  (b2e:counter-set start)
  (princ (strcat "\nСчётчик установлен в " (itoa start)
                 ". Следующий блок будет №" (itoa (1+ start))))
  (princ)
)

;; --- Выгрузка выбранных блоков ------------------------------------------

(defun c:BLOCK2EXCEL ( / ss i ename data rows filepath count)
  (princ "\n--- Извлечение координат блоков в Excel ---")
  (princ "\nВыберите блоки на чертеже...")
  (setq ss (ssget '((0 . "INSERT"))))
  (if (null ss)
    (princ "\nНичего не выбрано.")
    (progn
      (setq rows '() i 0 count (sslength ss))
      (while (< i count)
        (setq ename (ssname ss i)
              data  (b2e:get-block-data ename))
        (if data (setq rows (cons data rows)))
        (setq i (1+ i)))
      (setq rows (reverse rows))
      (princ (strcat "\nНайдено блоков: " (itoa (length rows))))
      (setq filepath (b2e:get-save-path))
      (if filepath
        (if (b2e:write-csv filepath rows)
          (b2e:open-in-excel filepath)))))
  (princ)
)

;; --- Выгрузка всех блоков чертежа ---------------------------------------

(defun c:BLOCK2EXCEL_ALL ( / ss i ename data rows filepath)
  (princ "\n--- Извлечение ВСЕХ блоков чертежа ---")
  (setq ss (ssget "_X" '((0 . "INSERT"))))
  (if (null ss)
    (princ "\nНа чертеже нет вхождений блоков.")
    (progn
      (setq rows '() i 0)
      (while (< i (sslength ss))
        (setq ename (ssname ss i)
              data  (b2e:get-block-data ename))
        (if data (setq rows (cons data rows)))
        (setq i (1+ i)))
      (setq rows (reverse rows))
      (princ (strcat "\nНайдено блоков: " (itoa (length rows))))
      (setq filepath (b2e:get-save-path))
      (if filepath
        (if (b2e:write-csv filepath rows)
          (b2e:open-in-excel filepath)))))
  (princ)
)

;; --- Выгрузка блоков по имени -------------------------------------------

(defun c:BLOCK2EXCEL_NAME ( / bname ss i ename data rows filepath)
  (princ "\n--- Извлечение блоков по имени ---")
  (setq bname (getstring T "\nВведите имя блока (можно с *, напр. POINT*): "))
  (if (or (null bname) (= bname ""))
    (princ "\nИмя не задано.")
    (progn
      (setq ss (ssget "_X" (list '(0 . "INSERT") (cons 2 bname))))
      (if (null ss)
        (princ (strcat "\nБлоки '" bname "' не найдены."))
        (progn
          (setq rows '() i 0)
          (while (< i (sslength ss))
            (setq ename (ssname ss i)
                  data  (b2e:get-block-data ename))
            (if data (setq rows (cons data rows)))
            (setq i (1+ i)))
          (setq rows (reverse rows))
          (princ (strcat "\nНайдено блоков: " (itoa (length rows))))
          (setq filepath (b2e:get-save-path))
          (if filepath
            (if (b2e:write-csv filepath rows)
              (b2e:open-in-excel filepath)))))))
  (princ)
)

;; ===========================================================================
;; DCL-ДИАЛОГИ (панель и настройки)
;; ===========================================================================

;; Создать временный DCL-файл
(defun b2e:make-dcl ( / dcl-path fp)
  (setq dcl-path (vl-filename-mktemp "b2e_dlg" nil ".dcl"))
  (setq fp (open dcl-path "w"))
  ;; --- Диалог "Панель" ---
  (write-line "b2e_panel : dialog {" fp)
  (write-line "  label = \"Block2Excel - панель управления\";" fp)
  (write-line "  : column {" fp)
  (write-line "    : boxed_column { label = \"Создание точек\";" fp)
  (write-line "      : button { key = \"btn_insert\"; label = \"Вставить точку с автономером\"; width = 40; }" fp)
  (write-line "      : button { key = \"btn_origin\"; label = \"Указать начальную точку\"; width = 40; }" fp)
  (write-line "      : button { key = \"btn_reset\";  label = \"Сбросить счётчик нумерации\"; width = 40; }" fp)
  (write-line "    }" fp)
  (write-line "    : boxed_column { label = \"Выгрузка в Excel\";" fp)
  (write-line "      : button { key = \"btn_sel\";  label = \"Выбрать блоки и выгрузить\"; width = 40; }" fp)
  (write-line "      : button { key = \"btn_all\";  label = \"Выгрузить ВСЕ блоки чертежа\"; width = 40; }" fp)
  (write-line "      : button { key = \"btn_name\"; label = \"Выгрузить блоки по имени...\"; width = 40; }" fp)
  (write-line "    }" fp)
  (write-line "    : boxed_column { label = \"Текущие параметры\";" fp)
  (write-line "      : text { key = \"lbl_origin\"; label = \"\"; }" fp)
  (write-line "      : text { key = \"lbl_corr\";   label = \"\"; }" fp)
  (write-line "      : text { key = \"lbl_scale\";  label = \"\"; }" fp)
  (write-line "      : text { key = \"lbl_count\";  label = \"\"; }" fp)
  (write-line "      : button { key = \"btn_settings\"; label = \"Настройки...\"; width = 40; }" fp)
  (write-line "    }" fp)
  (write-line "    : row {" fp)
  (write-line "      : button { key = \"accept\"; label = \"Закрыть\"; is_default = true; is_cancel = true; width = 15; }" fp)
  (write-line "    }" fp)
  (write-line "  }" fp)
  (write-line "}" fp)
  ;; --- Диалог "Настройки" ---
  (write-line "b2e_settings : dialog {" fp)
  (write-line "  label = \"Block2Excel - настройки\";" fp)
  (write-line "  : column {" fp)
  (write-line "    : boxed_column { label = \"Масштаб блоков\";" fp)
  (write-line "      : edit_box { key = \"scale\"; label = \"Масштаб (X=Y=Z): \"; edit_width = 12; }" fp)
  (write-line "      : text { label = \"Применяется к B2E_POINT и B2E_ORIGIN при вставке.\"; }" fp)
  (write-line "    }" fp)
  (write-line "    : boxed_column { label = \"Координаты начальной точки\";" fp)
  (write-line "      : edit_box { key = \"ox\"; label = \"X начальной точки: \"; edit_width = 18; }" fp)
  (write-line "      : edit_box { key = \"oy\"; label = \"Y начальной точки: \"; edit_width = 18; }" fp)
  (write-line "      : text { label = \"Автоматически обновляются командой 'Указать начальную точку'.\"; }" fp)
  (write-line "    }" fp)
  (write-line "    : boxed_column { label = \"Поправка к координатам начальной точки\";" fp)
  (write-line "      : edit_box { key = \"dx\"; label = \"dX (поправка по X): \"; edit_width = 18; }" fp)
  (write-line "      : edit_box { key = \"dy\"; label = \"dY (поправка по Y): \"; edit_width = 18; }" fp)
  (write-line "      : text { label = \"Формула пересчёта при выгрузке:\"; }" fp)
  (write-line "      : text { label = \"  X_итог = X_нач - dX - X_точки\"; }" fp)
  (write-line "      : text { label = \"  Y_итог = Y_нач - dY - Y_точки\"; }" fp)
  (write-line "    }" fp)
  (write-line "    : row {" fp)
  (write-line "      : button { key = \"accept\"; label = \"OK\"; is_default = true; width = 12; }" fp)
  (write-line "      : button { key = \"cancel\"; label = \"Отмена\"; is_cancel = true; width = 12; }" fp)
  (write-line "    }" fp)
  (write-line "  }" fp)
  (write-line "}" fp)
  (close fp)
  dcl-path
)

;; Диалог настроек
(defun b2e:show-settings ( / dcl-id dcl-path ret v)
  (setq dcl-path (b2e:make-dcl))
  (setq dcl-id (load_dialog dcl-path))
  (if (not (new_dialog "b2e_settings" dcl-id))
    (progn (unload_dialog dcl-id) (princ "\nОшибка диалога настроек.") nil)
    (progn
      (set_tile "scale" (rtos (b2e:get-scale)    2 4))
      (set_tile "ox"    (rtos (b2e:get-origin-x) 2 4))
      (set_tile "oy"    (rtos (b2e:get-origin-y) 2 4))
      (set_tile "dx"    (rtos (b2e:get-corr-x)   2 4))
      (set_tile "dy"    (rtos (b2e:get-corr-y)   2 4))
      (action_tile "accept"
        (strcat "(setq *b2e:tmp* "
                "(list (atof (get_tile \"scale\")) "
                "      (atof (get_tile \"ox\")) "
                "      (atof (get_tile \"oy\")) "
                "      (atof (get_tile \"dx\")) "
                "      (atof (get_tile \"dy\")))) "
                "(done_dialog 1)"))
      (action_tile "cancel" "(done_dialog 0)")
      (setq ret (start_dialog))
      (unload_dialog dcl-id)
      (vl-file-delete dcl-path)
      (if (and (= ret 1) *b2e:tmp*)
        (progn
          (setq v *b2e:tmp*)
          (if (<= (nth 0 v) 0.0)
            (progn (princ "\nМасштаб должен быть > 0. Установлено 1.0.")
                   (b2e:set-scale 1.0))
            (b2e:set-scale (nth 0 v)))
          (b2e:set-origin (nth 1 v) (nth 2 v))
          (b2e:set-corr   (nth 3 v) (nth 4 v))
          (setq *b2e:tmp* nil)
          (princ "\nНастройки сохранены.")
          T)
        nil))
  )
)

;; --- Главная панель ------------------------------------------------------

(defun c:B2E ( / dcl-id dcl-path ret action)
  (b2e:ensure-point-block)
  (b2e:ensure-origin-block)
  (setq action nil)
  (setq dcl-path (b2e:make-dcl))
  (setq dcl-id (load_dialog dcl-path))
  (if (not (new_dialog "b2e_panel" dcl-id))
    (progn (unload_dialog dcl-id) (princ "\nОшибка диалога."))
    (progn
      (set_tile "lbl_origin"
        (strcat "Нач. точка: X=" (rtos (b2e:get-origin-x) 2 3)
                "  Y=" (rtos (b2e:get-origin-y) 2 3)))
      (set_tile "lbl_corr"
        (strcat "Поправка:   dX=" (rtos (b2e:get-corr-x) 2 3)
                "  dY=" (rtos (b2e:get-corr-y) 2 3)))
      (set_tile "lbl_scale"
        (strcat "Масштаб блоков: " (rtos (b2e:get-scale) 2 3)))
      (set_tile "lbl_count"
        (strcat "Счётчик нумерации: " (itoa (b2e:counter-get))))
      (action_tile "btn_insert"   "(setq *b2e:act* \"INSERT\")   (done_dialog 2)")
      (action_tile "btn_origin"   "(setq *b2e:act* \"ORIGIN\")   (done_dialog 2)")
      (action_tile "btn_reset"    "(setq *b2e:act* \"RESET\")    (done_dialog 2)")
      (action_tile "btn_sel"      "(setq *b2e:act* \"SEL\")      (done_dialog 2)")
      (action_tile "btn_all"      "(setq *b2e:act* \"ALL\")      (done_dialog 2)")
      (action_tile "btn_name"     "(setq *b2e:act* \"NAME\")     (done_dialog 2)")
      (action_tile "btn_settings" "(setq *b2e:act* \"SETTINGS\") (done_dialog 2)")
      (action_tile "accept"       "(setq *b2e:act* nil) (done_dialog 0)")
      (setq ret (start_dialog))
      (unload_dialog dcl-id)
      (vl-file-delete dcl-path)
      (setq action *b2e:act* *b2e:act* nil)
      (cond
        ((= action "INSERT")   (c:B2E_INSERT))
        ((= action "ORIGIN")   (c:B2E_ORIGIN))
        ((= action "RESET")    (c:B2E_RESET))
        ((= action "SEL")      (c:BLOCK2EXCEL))
        ((= action "ALL")      (c:BLOCK2EXCEL_ALL))
        ((= action "NAME")     (c:BLOCK2EXCEL_NAME))
        ((= action "SETTINGS")
         (b2e:show-settings)
         (c:B2E)))                ; повторно открываем панель
    ))
  (princ)
)

(defun c:B2E_SETTINGS ( / )
  (b2e:show-settings)
  (princ)
)

;; ===========================================================================
;; УСТАНОВКА ПАНЕЛИ ИНСТРУМЕНТОВ AutoCAD
;; ===========================================================================
;; Создаёт классическую плавающую панель Toolbar через AutoCAD COM API.
;; Панель появляется сразу после выполнения команды и сохраняется в acad.cuix
;; пользователя - доступна во всех чертежах после установки.
;;
;; Команды:
;;   B2E_INSTALL_PANEL   - установить (создать) панель Block2Excel
;;   B2E_UNINSTALL_PANEL - удалить панель из интерфейса
;; ===========================================================================

(setq *b2e:panel-name* "Block2Excel")

;; Безопасное добавление кнопки с обработкой ошибок
(defun b2e:add-button (toolbar idx name macro / btn)
  (setq btn
    (vl-catch-all-apply
      (function (lambda ()
        (vla-AddToolbarButton toolbar idx name name macro)))))
  (if (vl-catch-all-error-p btn)
    (progn
      (princ (strcat "\n  ! Не удалось добавить '" name "': "
                     (vl-catch-all-error-message btn)))
      nil)
    btn)
)

;; Безопасное добавление разделителя
(defun b2e:add-separator (toolbar idx / sep)
  (vl-catch-all-apply
    (function (lambda ()
      (vla-AddSeparator toolbar idx))))
)

(defun c:B2E_INSTALL_PANEL ( / acad menugroups mg toolbars tb existing)
  (vl-load-com)
  (princ "\n--- Установка панели инструментов Block2Excel ---")
  (setq acad (vlax-get-acad-object))
  (setq menugroups (vla-get-MenuGroups acad))

  ;; Берём главную группу меню (обычно ACAD)
  (setq mg
    (vl-catch-all-apply
      (function (lambda () (vla-Item menugroups 0)))))
  (if (vl-catch-all-error-p mg)
    (progn (princ "\nОшибка доступа к меню AutoCAD.") (exit)))

  (setq toolbars (vla-get-Toolbars mg))

  ;; Если панель уже есть - удалим её, чтобы пересоздать с актуальными командами
  (setq existing
    (vl-catch-all-apply
      (function (lambda () (vla-Item toolbars *b2e:panel-name*)))))
  (if (not (vl-catch-all-error-p existing))
    (progn
      (princ "\nПанель уже существует, пересоздаю...")
      (vla-Delete existing)))

  ;; Создаём новую панель
  (princ (strcat "\nСоздаю панель '" *b2e:panel-name* "'..."))
  (setq tb (vla-Add toolbars *b2e:panel-name*))

  ;; --- Кнопки -----------------------------------------------------------
  ;; Группа 1: создание точек
  (b2e:add-button tb 0 "B2E - Панель"          "^C^C_B2E ")
  (b2e:add-separator tb 1)
  (b2e:add-button tb 2 "Вставить точку"        "^C^C_B2E_INSERT ")
  (b2e:add-button tb 3 "Начальная точка"       "^C^C_B2E_ORIGIN ")
  (b2e:add-button tb 4 "Сброс счётчика"        "^C^C_B2E_RESET ")
  (b2e:add-separator tb 5)
  ;; Группа 2: выгрузка
  (b2e:add-button tb 6 "Выгрузить выбранные"   "^C^C_BLOCK2EXCEL ")
  (b2e:add-button tb 7 "Выгрузить все"         "^C^C_BLOCK2EXCEL_ALL ")
  (b2e:add-button tb 8 "Выгрузить по имени"    "^C^C_BLOCK2EXCEL_NAME ")
  (b2e:add-separator tb 9)
  ;; Группа 3: настройки
  (b2e:add-button tb 10 "Настройки"            "^C^C_B2E_SETTINGS ")

  ;; Показать панель
  (vla-put-Visible tb :vlax-true)
  (vla-put-FloatingRows tb 1)

  ;; Сохранить настройки меню в acad.cuix пользователя
  (vl-catch-all-apply
    (function (lambda ()
      (vla-Save mg acMenuFileCompiled))))

  (princ (strcat "\nГотово. Панель '" *b2e:panel-name*
                 "' установлена и видна."))
  (princ "\nДля сохранения положения панели в рабочем пространстве:")
  (princ "\n  1. Передвиньте панель в удобное место")
  (princ "\n  2. Выполните команду WSSAVE и сохраните текущее РП")
  (princ)
)

(defun c:B2E_UNINSTALL_PANEL ( / acad menugroups mg toolbars existing)
  (vl-load-com)
  (setq acad (vlax-get-acad-object))
  (setq menugroups (vla-get-MenuGroups acad))
  (setq mg
    (vl-catch-all-apply
      (function (lambda () (vla-Item menugroups 0)))))
  (if (vl-catch-all-error-p mg)
    (progn (princ "\nОшибка доступа к меню.") (exit)))
  (setq toolbars (vla-get-Toolbars mg))
  (setq existing
    (vl-catch-all-apply
      (function (lambda () (vla-Item toolbars *b2e:panel-name*)))))
  (if (vl-catch-all-error-p existing)
    (princ (strcat "\nПанель '" *b2e:panel-name* "' не найдена."))
    (progn
      (vla-Delete existing)
      (vl-catch-all-apply
        (function (lambda () (vla-Save mg acMenuFileCompiled))))
      (princ (strcat "\nПанель '" *b2e:panel-name* "' удалена."))))
  (princ)
)

;; ===========================================================================
;; ИНИЦИАЛИЗАЦИЯ ПРИ ЗАГРУЗКЕ
;; ===========================================================================

(b2e:init-defaults)
(vl-catch-all-apply 'b2e:ensure-point-block)
(vl-catch-all-apply 'b2e:ensure-origin-block)

(princ "\n=== BLOCK2EXCEL загружен ===")
(princ "\nГлавная команда:")
(princ "\n  B2E                  - открыть панель управления")
(princ "\nПанель инструментов:")
(princ "\n  B2E_INSTALL_PANEL    - установить панель в интерфейс AutoCAD")
(princ "\n  B2E_UNINSTALL_PANEL  - удалить панель")
(princ "\nОтдельные команды:")
(princ "\n  B2E_INSERT           - вставить блок с автономером")
(princ "\n  B2E_ORIGIN           - указать начальную точку")
(princ "\n  B2E_RESET            - сбросить счётчик нумерации")
(princ "\n  B2E_SETTINGS         - настройки (масштаб, поправки)")
(princ "\n  BLOCK2EXCEL          - выбрать блоки и выгрузить")
(princ "\n  BLOCK2EXCEL_ALL      - выгрузить все блоки")
(princ "\n  BLOCK2EXCEL_NAME     - выгрузить блоки по имени")
(princ (strcat "\nМасштаб: " (rtos (b2e:get-scale) 2 3)
               "  Счётчик: " (itoa (b2e:counter-get))
               "  Нач.точка: ("
               (rtos (b2e:get-origin-x) 2 2) ", "
               (rtos (b2e:get-origin-y) 2 2) ")"))
(princ "\nДля начала работы: B2E   |   Установить панель: B2E_INSTALL_PANEL")
(princ)
