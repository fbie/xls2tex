#lang racket

(require xml)

;; Works up to column 26. I am too lazy to implement a general
;; solution, and how likely is it that our spreadsheets exceed 26
;; columns?
;; TODO: Make work for columns > 26.
(define (xls/c1->column col)
    (string (integer->char (+ 64 col))))

(struct cellref (col row cabs rabs) #:transparent)

(define (xls/cellref-row-abs? ref)
  (cellref-rabs ref))

(define (xls/cellref-col-abs? ref)
  (cellref-cabs ref))

(struct pos (row col) #:transparent)

;; Always start at index (1, 1).
(define xls-start-pos (pos 1 1))

;; Format for R1C1. This is easy, we only need to decide on the
;; formatting string.
(define (xls/print-r1c1 ref)
  (let [(r (format (if (xls/cellref-row-abs? ref) "R~a" "R[~a]") (cellref-row ref)))
        (c (format (if (xls/cellref-col-abs? ref) "C~a" "C[~a]") (cellref-col ref)))]
    (string-join (list r c) "")))

;; Format for A1. This is more complicated, because we need to
;; compute relative references wrt. the current position.
(define (xls/print-a1 ref pos)
  (let [(r (if (xls/cellref-row-abs? ref)
               (format "$~a" (cellref-row ref))
               (number->string (+ (cellref-row ref) (pos-row pos)))))
        (c (if (xls/cellref-col-abs? ref)
               (format "$~a" (xls/c1->column (cellref-col ref)))
               (xls/c1->column (+ (cellref-col ref) (pos-col pos)))))]
    (string-join (list c r) "")))

;; Print a cell reference using either 'R1C1 or 'A1 styles. If 'A1 is
;; desired, you need to pass in a valid position, too.
(define (xls/cellref-print style ref pos)
  (match style
    ['R1C1 (xls/print-r1c1 ref)]
    ['A1 (xls/print-a1 ref pos)]))

;; Parse a R1C1 formatted cell reference and return a r1c1ref struct.
(define (xls/parse-r1c1 refstring)
  (let [(get-num (λ (s) (string->number (or (car (regexp-match #px"[-]?\\d+" s)) "0"))))
        (row-abs (regexp-match #px"R\\d+" refstring))
        (row-rel (regexp-match #px"R\\[[+-]\\d+\\]" refstring))
        (col-abs (regexp-match #px"C\\d+" refstring))
        (col-rel (regexp-match #px"C\\[[+-]\\d+\\]" refstring))]
    (cellref (get-num (car (or col-rel col-abs)))
             (get-num (car (or row-rel row-abs)))
             col-abs
             row-abs)))

;; Convert a R1C1 formatted refstring to an A1 formatted refstring
;; using pos.
(define (xls/r1c1->a1 refstring pos)
  (xls/cellref-print 'A1 (xls/parse-r1c1 refstring) pos))

(define (xls/pos-advance-col curr n)
  (pos (pos-row curr) (+ n (pos-col curr))))

;; Advance column by one.
(define (xls/pos-next-col curr)
  (xls/pos-advance-col curr 1))

;; Advance row by one.
(define (xls/pos-next-row curr)
  (pos (+ 1 (pos-row curr)) 1))

;; Filter elements recursively from a list xs using the given predicate.
(define (xls/filter-recursive predicate lst)
  (define (filter-recurse predicate lst)
    (if (list? lst)
        (filter predicate (map (curry xls/filter-recursive predicate) lst))
        lst))
  ;; Wrap predicate and include a check for empty lists, we don't want
  ;; to keep those.
  (filter-recurse (λ (x) (and (not (empty? x)) (predicate x))) lst))

;; Unwrap singleton lists to their single element. Does nothing to
;; empty lists.
(define (xls/unwrap lst)
  (if (= (length lst) 1)
      (car lst)
      lst))

;; Unwrap lists recursively.
(define (xls/flatten lst)
  (if (list? lst)
      (xls/unwrap (map xls/flatten lst))
      lst))

;; Read the XLS formatted file into an xexpr.
;; TODO: Handle non UTF-8 smartly.
;; TODO: Remove newline returns.
(define (xls/read filename)
  (let [(predicate (λ (x)
                     (not (and (string? x)
                               (regexp-match #px"\r\n\\s*" x)))))
        (doc (xml->xexpr (document-element (with-input-from-file filename read-xml))))]
    (xls/flatten (xls/filter-recursive predicate doc))))

;; A variant of assoc that is more robust to non-pairs in lists.
(define (xls/xexpr-assoc key dict)
  (if (empty? dict)
      #f
      (let [(candidate (car dict))]
        (if (and (list? candidate) (eq? (car candidate) key))
            candidate
            (xls/xexpr-assoc key (cdr dict))))))

(define (xls/get key dict)
  (let [(val (xls/xexpr-assoc key dict))]
    (if val (car (cdr val)) #f)))

;; Retrieve a formula from an element, if any.
(define (xls/get-formula elem)
  (xls/get 'ss:Formula (car elem)))

;; Retrieve an index offset from an element, if any.
(define (xls/get-index elem)
  (let [(idx (xls/get 'ss:Index (car elem)))]
    (if idx
        (string->number idx)
        1)))

;; Retrieve datra from an element, if any.
(define (xls/get-data elem)
  (car (cdr (cdr (xls/xexpr-assoc 'Data elem)))))

;; Take all pairs from reps and apply them to str using
;; string-replace. The result is str with all occurrences of the first
;; element replaced with the second element for all peirs in reps.
(define (xls/replace-all reps str)
  (foldr (λ (rep str) (string-replace str (car rep) (cdr rep))) str reps))

;; Convert an R1C1 formula into an A1 formula for some absolute
;; position.
(define (xls/r1c1formula->a1formula formula pos)
  (let* [(refs (regexp-match* #px"R\\[?[+-]\\d+\\]?C\\[?-?\\d+\\]?" formula))
         (reps (map (λ (ref) (cons ref (xls/r1c1->a1 ref pos))) refs))]
    (xls/replace-all reps formula)))

;; Return a string that represents a cell in TeX table format.
(define (xls/cell->tex cell pos)
  (let* [(ccell (cdr cell))
         (formula (xls/get-formula ccell))
         (index (xls/get-index ccell)) ;; TODO: Meh.
         (data (xls/get-data ccell))
         (npos (xls/pos-advance-col pos index))]
    (cons
     (format "~a \\texttt{~a}" ;; Only for skipping rows.
             (string-append* (make-list index " &"))
             (or (if formula (xls/r1c1formula->a1formula formula npos) #f) data))
     npos)))

(define (xls/filter-type type lst)
  (filter (λ (elem) (and (list? elem) (eq? (car elem) type))) lst))

;; Return a string that represents a row in TeX table format
(define (xls/row->tex row pos)
  (cons
   (foldl (λ (cell state)
            (let [(res (xls/cell->tex cell (cdr state)))]
              (cons
               (format "~a ~a" (car state) (car res))
               (cdr res))))
          (cons (number->string (pos-row pos)) pos)
          (xls/filter-type 'Cell row))
   (xls/pos-next-row pos)))

(define (xls/get-name elem)
  (xls/get 'ss:Name elem))

;; Get the first sheet with the correct sheet name. There should not
;; be two sheets with the same name.
(define (xls/get-sheet sheetname doc)
  (car (filter (λ (sheet)
                 (string=? sheetname (xls/get-name sheet)))
               (xls/filter-type 'Worksheet doc))))

(define (xls/sheet->tex sheet doc)
  (let [(rows (xls/filter-type 'Row (xls/xexpr-assoc 'Table (xls/get-sheet sheet doc))))]
    (foldl (λ (row state)
             (let [(res (xls/row->tex row (cdr state)))]
               (cons
                (format "~a ~a\\\\ \\hline \n" (car state) (car res))
                (cdr res))))
           (cons "##TESTHEADER \\\\ \\hline \n" xls-start-pos)
           rows)))

(define (xls/doc->tex doc)
  (xls/filter-type 'Worksheet doc))

;; TODO: For testing, delete later.
(define doc (xls/read "test.xml"))
(define rows (xls/sheet->tex "Sheet1" doc))
