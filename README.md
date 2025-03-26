# gitstudy
study

=HYPERLINK(
    LEFT(A1, FIND("[", A1)-1) & 
    MID(A1, FIND("[", A1)+1, FIND("]", A1)-FIND("[", A1)-1) & 
    "#" & 
    MID(A1, FIND("]", A1)+1, FIND("!", A1)-FIND("]", A1)-1) & 
    "!" & 
    SUBSTITUTE(RIGHT(A1, LEN(A1)-FIND("!", A1)), "$", ""),
    "Open File & Jump"
)
