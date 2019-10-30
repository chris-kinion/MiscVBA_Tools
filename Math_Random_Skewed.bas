Attribute VB_Name = "Math_Random_Skewed"
'Author:  Stephen H.Gersuk, 2008

Function RandSkew(fAlpha As Single, _
                  Optional fLocation As Single = 0!, _
                  Optional fScale As Single = 1!, _
                  Optional bVolatile As Boolean = False) As Single

    ' shg 2008-0919
    ' http://azzalini.stat.unipd.it/SN/faq.html         (algorithm)
    ' http://azzalini.stat.unipd.it/SN/Intro/intro.html (intro)
    ' http://azzalini.stat.unipd.it/SN/plot-SN1.html    (density function plotting)

    ' Returns a random variable with skewed distribution
    '       fAlpha      = skew or 'shape'
    '       fLocation   = location
    '       fScale > 0  = scale

    Dim fSigma  As Single  ' correlation coefficient derived from alpha
    Dim afRN()  As Single  ' pair of random normal variates
    Dim u0      As Single  ' see algorithm
    Dim v       As Single  ' see algorithm
    Dim u1      As Single  ' see algorithm

    If bVolatile Then Application.Volatile
    Randomize (Timer)

    fSigma = fAlpha / Sqr(1 + fAlpha * fAlpha)

    afRN = RandNorm()
    u0 = afRN(1)
    v = afRN(2)
    u1 = fSigma * u0 + Sqr(1! - fSigma * fSigma) * v

    RandSkew = IIf(u0 >= 0, u1, -u1) * fScale + fLocation
End Function

Function RandNorm(Optional Mean As Single = 0!, _
                  Optional Dev As Single = 1!, _
                  Optional bVolatile As Boolean = False) As Single()
    ' shg 1999-1103
    ' Returns a pair of random deviates (Singles) with the specified
    ' mean and deviation. Orders of magnitude faster than
    '     =NORMINV(RAND(), Mean, Dev)

    ' Box-Muller Polar Method
    ' Donald Knuth, The Art of Computer Programming,
    ' Vol 2, Seminumerical Algorithms, p. 117

    Dim af(1 To 2) As Single
    Dim x       As Single
    Dim y       As Single
    Dim w       As Single

    If bVolatile Then Application.Volatile

    Do
        x = 2! * Rnd - 1!
        y = 2! * Rnd - 1!
        w = x ^ 2 + y ^ 2
    Loop Until w < 1!

    w = Sqr((-2! * CSng(Log(w))) / w)
    af(1) = Dev * x * w + Mean
    af(2) = Dev * y * w + Mean
    RandNorm = af
End Function

