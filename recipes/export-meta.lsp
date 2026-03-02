;; export-meta.lsp
;;
;; This AutoLISP routine exports all metadata for every object in a DWG file
;; to a JSON file. It is designed to be run from the AutoCAD Core Console.

;; Helper function to escape special JSON characters in a string.
(defun escape-json-string (str)
  (setq str (vl-string-subst "\\\\" "\\" str))   ; Escape backslashes
  (setq str (vl-string-subst "\\\"" "\"" str))   ; Escape double quotes
  (setq str (vl-string-subst "\\n" "\n" str))     ; Escape newlines
  (setq str (vl-string-subst "\\r" "\r" str))     ; Escape carriage returns
  (setq str (vl-string-subst "\\t" "\t" str))     ; Escape tabs
  str
)

;; Helper function to format a DXF value for JSON output.
(defun format-dxf-value (val)
  (cond
    ((= (type val) 'STR) (strcat "\"" (escape-json-string val) "\"")) ; String
    ((= (type val) 'REAL) (rtos val 2 8))                           ; Real number
    ((= (type val) 'INT) (itoa val))                                 ; Integer
    ((= (type val) 'LIST)                                            ; 2D/3D Point
      (strcat "[" (rtos (car val) 2 8) "," (rtos (cadr val) 2 8)
        (if (caddr val) (strcat "," (rtos (caddr val) 2 8)) "") "]"
      )
    )
    ((= (type val) 'ENAME)                                           ; Entity Name (reference)
      (strcat "\"<Entity Name: " (vl-princ-to-string val) ">\"")
    )
    (t (strcat "\"" (vl-princ-to-string val) "\""))                   ; Any other type
  )
)

;; Main function to be called from the script.
(defun C:EXPORTMETATOJSON ( / dwg_path dwg_name json_path outfile first_obj ent edata)
  (princ "\nStarting metadata export to JSON...")

  ;; --- Setup File Paths ---
  (setq dwg_path (getvar "DWGPREFIX"))
  (setq dwg_name (vl-filename-base (getvar "DWGNAME")))
  (setq json_path (strcat dwg_path dwg_name "_metadata.json"))

  ;; --- Open File for Writing ---
  (setq outfile (open json_path "w"))
  (if (not outfile)
    (progn
      (princ (strcat "\nError: Could not open file for writing at " json_path))
      (exit)
    )
  )

  ;; --- Write JSON Header ---
  (write-line "{" outfile)
  (write-line (strcat "  \"drawingName\": \"" (getvar "DWGNAME") "\",") outfile)
  (write-line (strcat "  \"exportDate\": \"" (rtos (getvar "CDATE") 2 8) "\",") outfile)
  (write-line "  \"objects\": [" outfile)

  ;; --- Iterate Through All Entities and Write to File ---
  (setq first_obj T) ; Flag to handle comma separation
  (setq ent (entnext)) ; Get the first entity in the database

  (while ent
    (if first_obj
      (setq first_obj nil)
      (write-line "," outfile) ; Add a comma before each new object (except the first)
    )

    (setq edata (entget ent)) ; Get the entity's DXF data list

    (write-line "    {" outfile)
    ; Write the handle (DXF code 5)
    (write-line (strcat "      \"handle\": \"" (cdr (assoc 5 edata)) "\",") outfile)
    ; Write the entity type (DXF code 0)
    (write-line (strcat "      \"type\": \"" (cdr (assoc 0 edata)) "\",") outfile)
    ; Write all DXF data
    (write-line "      \"dxfData\": [" outfile)

    ;; Loop through each DXF group code pair (e.g., '(0 . "LINE"))
    (foreach pair edata
      (write-line
        (strcat "        { \"code\": " (itoa (car pair)) ", \"value\": " (format-dxf-value (cdr pair)) " }"
          (if (not (equal pair (last edata))) ",") ; Add comma if not the last item
        )
        outfile
      )
    )

    (write-line "      ]" outfile) ; Close dxfData array
    (write-line "    }" outfile)   ; Close object

    (setq ent (entnext ent)) ; Get the next entity
  )

  ;; --- Write JSON Footer and Close File ---
  (write-line "\n  ]" outfile) ; Close objects array
  (write-line "}" outfile)   ; Close main JSON object
  (close outfile)

  (princ (strcat "\nExport complete. File saved to: " json_path))
  (princ) ; Suppress returning the last value
)