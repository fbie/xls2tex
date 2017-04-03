#lang racket

(require xml)

;;; Construction of a usable row- and cell-representation.

(struct pos (row col)
  #:guard (λ (row col name)
            (unless (and (number? row) (number? col))
              (error "Not a valid pos."))
            (values row col))
  #:transparent)

(define zero (pos 0 0))

;; Advance column by one or set to n.
(define (xls/pos-next-col curr [n #f])
  (if n
      (pos (pos-row curr) (string->number n))
      (pos (pos-row curr) (add1 (pos-col curr)))))

;; Advance row by one or set to n.
(define (xls/pos-next-row curr [n #f])
  (if n
      (pos (string->number n) (pos-col curr))
      (pos (add1 (pos-row curr)) (pos-col curr))))

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
  (if (or (not (list? dict)) (empty? dict))
      #f
      (let [(candidate (car dict))]
        (if (and (list? candidate) (eq? (car candidate) key))
            candidate
            (xls/xexpr-assoc key (cdr dict))))))

;; Get value for some key from an xls structure. Returns #f if the key
;; doesn't exist.
(define (xls/get key dict)
  (let [(val (xls/xexpr-assoc key dict))]
    (if val
        (cadr val)
        #f)))

;; Retrieve a formula from an element, if any.
(define (xls/get-formula elem)
  (xls/get 'ss:Formula (car elem)))

;; Retrieve an index offset from an element.
(define (xls/get-index elem)
  (or (xls/get 'ss:Index elem) (xls/get 'ss:Index (car elem))))

;; Retrieve data from an element, if any.
(define (xls/get-data elem)
  (caddr (xls/xexpr-assoc 'Data elem)))

;; Retrieve the element's name tag.
(define (xls/get-name elem)
  (xls/get 'ss:Name elem))

;; A cell consists of an expression and a position.
;; TODO: Maybe only use column?
(struct cell (expr pos)
  #:guard (λ (expr pos name)
            (unless (and (string? expr) (pos? pos))
                (error "Not a valid cell."))
              (values expr pos))
  #:transparent)

;; Convert a cell to a TeX string.
(define (cell->tex cell refstyle)
  (match refstyle
    ['R1C1 (cell-expr cell)]
    ['A1 (xls/r1c1formula->a1formula (cell-expr cell) (cell-pos cell))]))

;; Return a string that represents a cell in the spreadsheet.
(define (xls/parse-cell c p)
  (let* [(ccell (cdr c))
         (formula (xls/get-formula ccell))
         (index (xls/get-index ccell))
         (data (xls/get-data ccell))]
    (cell (or formula data) (xls/pos-next-col p index))))

;; Retrieve all elements from the xls that are of the given type.
(define (xls/filter-type type lst)
  (filter (λ (elem) (and (list? elem) (eq? (car elem) type))) lst))

;; Typical prefix-sum. If lst has n elements, then the result has n +
;; 1, where the first element is the original prefix.
(define (scan f pre lst)
  (match (length lst)
    [0 (list pre)]
    [_ (cons pre (scan f (f pre (car lst)) (cdr lst)))]))

;; Return a string that represents a row in TeX table format together
;; with the last position.
(define (xls/row->cells row pos)
  (unless (list? row)
    (error "Row must be a list!"))
  (cdr (scan (λ (p c)
               (unless (cell? p)
                 (error "Prefix not a cell."))
               (xls/parse-cell c (cell-pos p)))
             (cell "" pos) ;; Fake cell that we strip later.
             row)))

;; Get the first sheet with the correct sheet name. There should not
;; be two sheets with the same name.
(define (xls/get-sheet sheetname doc)
  (car (filter (λ (sheet)
                 (string=? sheetname (xls/get-name sheet)))
               (xls/filter-type 'Worksheet doc))))

(define (xls/sheet->cells sheet doc)
  (let* [(rows (xls/filter-type 'Row (xls/xexpr-assoc 'Table (xls/get-sheet sheet doc))))
         ;; Compute row numbers; allows us to use map instead of scan.
         (irows (cdr (scan (λ (pre row) (cons (xls/pos-next-row (car pre) (xls/get-index row)) row))
                           (cons zero empty)
                           rows)))]
    (for/list ([irow irows])
      (let [(pos (car irow))
            (row (cdr irow))]
        (xls/row->cells (xls/filter-type 'Cell row) pos)))))

;; TODO: For testing, delete later.
(define doc (xls/read "test.xml"))
(define rows (xls/sheet->cells "Sheet1" doc))


;;; Cell reference formatting.

;; Works up to column 26. I am too lazy to implement a general
;; solution, and how likely is it that our spreadsheets exceed 26
;; columns?
;; TODO: Make work for columns > 26.
(define (xls/c1->column col)
    (string (integer->char (+ 64 col))))

(struct cellref (col row cabs rabs)
  #:guard (λ (col row cabs rabs name)
            (unless (and (number? col) (number? row) (boolean? cabs) (boolean? rabs))
              (error "Not a valid cellref."))
            (values col row cabs rabs))
  #:transparent)

(define (xls/cellref-row-abs? ref)
  (cellref-rabs ref))

(define (xls/cellref-col-abs? ref)
  (cellref-cabs ref))

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
(define (xls/cellref-print style ref [pos zero])
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

;; Take all pairs from reps and apply them to str using
;; string-replace. The result is str with all occurrences of the first
;; element replaced with the second element for all pairs in reps.
(define (xls/replace-all reps str)
  (foldr (λ (rep str) (string-replace str (car rep) (cdr rep))) str reps))

;; Convert an R1C1 formula into an A1 formula for some absolute
;; position.
(define (xls/r1c1formula->a1formula formula pos)
  (let* [(refs (regexp-match* #px"R\\[?[+-]\\d+\\]?C\\[?-?\\d+\\]?" formula))
         (reps (map (λ (ref) (cons ref (xls/r1c1->a1 ref pos))) refs))]
    (xls/replace-all reps formula)))
