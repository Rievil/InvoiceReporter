Imports Microsoft.VisualBasic

Public Class Reporter
    '-----------------------------------
    'Poslední úprava 15.09.2019 13:21
    'Úprava pro více listů v 1 soubou
    '--úprava pro VCP
    '-----------------------------------
    'H:\Google drive\GitHub\InvoiceReporter\Reporter\ReporterClass.vb
    'integer______________________________________
    Public suma_polozek As Integer 'pocet zjistenych polozek ze spp
    Public pocet_fakturovanych_polozek As Integer
    Public nacteny_rozmery As Integer
    Public nacteny_radek As Integer
    Public nacteny_sloupec As Integer
    Public start_pocitani As Date
    Public konec_pocitani As Date
    Public obdobi_fakturace As String
    Public stranka_protokolu As Integer
    Public aktiv_list As String
    Public aktiv_barva As String
    Public aktiv_so As String
    Public SkutecnyList As Integer
    Public cislo_listu As Integer
    Public CelkovyPocetListu As Integer
    Public PocetObsazenychRadku As Integer
    'integer____promene nastevni makra____________
    Public yZacatekPolozekSPP As Integer
    Public xAktivSo As Integer
    Public yAktivSo As Integer
    Public xCisloListu As Integer
    Public yCisloListu As Integer
    Public xSloupecPolozek As Integer
    Public xSloupecZustatekMnozstvi As Integer
    Public xSloupecMJ As Integer
    Public yStartRadek As Integer
    Public xProtAktivSo As Integer
    Public yProtAktivSo As Integer
    Public xCisloSloupcePolozek As Integer
    Public xSloupecMnozstviVPoli As Integer
    Public xSloupecSledovaneObdobi As Integer
    Public xSloupecHistorieFakturace As Integer
    Public xSloupecSVyhledat As Integer
    Public xSloupecPopisPolozky As Integer
    Public NazevListSPP As String
    Public xSloupecSVyhledatZustatek As Integer
    Public NadpisKontrolaZustatek As String
    Public xSloupecMJHlavicka As Integer
    Dim xSloupecKonecPopisu As Integer
    Public xSloupecOdkazPopis As Integer 'slopes odkazu tabuly v SVYHLEDAT s popisem
    Public xSloupecOdkazMJ As Integer 'slopes odkazu tabuly v SVYHLEDAT s mernou jednotkou
    Public xSloupecOdkazZustatek As Integer  'slopes odkazu tabuly v SVYHLEDAT s zustatkem
    'string_______________________________________

    'pole_________________________________________
    Public polozka(0 To 500, 0 To 6) As Object
    Public pol_list(0 To 500) As String
    Public mem_pol(0 To 500, 0 To 4) As Object
    Public Radek(1 To 100) As Object
    Public Sloupec(1 To 100) As Object
    Public znak(1 To 10) As String
    Public pokam_zformatovat(0 To 300) As Integer
    Public List(0 To 2000) As Integer

    Sub cbSectiPolozky_click()
        Application.ScreenUpdating = False
        'nahraje zakladni promene
        'souradnic jednotlivych bunek
        DefaultVariables()

        cislo_listu = Empty
        Erase List

        nacteny_radek = 1
        nacteny_sloupec = 1

        aktiv_list = ActiveSheet.Name
        aktiv_barva = ActiveSheet.Tab.Color
        'uprava zjisteni nazvu SO
        aktiv_so = Cells(yAktivSo, xAktivSo)
        aktiv_so = Replace(aktiv_so, " ", "")

        nacti_spp_polozky()
        projdi_listy()
        vypis_soucty_do_spp()
        dopln_celk_cislo_protokolu()

        Application.ScreenUpdating = True

        MsgBox("Pocet fakturovanych polozek: " & pocet_fakturovanych_polozek & " polozek/y") ' & vbCrLf & "Pocitani trvalo: " & ubehly_cas
    End Sub

    Function DefaultVariables()
        'Soupis polozek-----------------------------------
        yZacatekPolozekSPP = 13 'radke od ktereho zacinaji polozky v SPP
        xAktivSo = 3    'radek bunky ve kterem je nazev stavebniho objektu na liste SPP
        yAktivSo = 6    'sloupec bunky ve kterem je nazev stavebniho objektu na liste SPP
        xSloupecPolozek = 1 'sloupec ve kterem jsou polozky na SPP
        xSloupecZustatekMnozstvi = Alpha2Number("O")   'sloupec ve kterem je zustatkove mnozstvi v SPP
        xSloupecMJ = Alpha2Number("D")                  'sloupec ve kterem jsou merne jednotky v SPP
        xSloupecSledovaneObdobi = Alpha2Number("K")    'sloupec s mnozstvim sledovaneho obdobi v SPP
        xSloupecHistorieFakturace = Alpha2Number("W")  'sloupec do ktereho se ma nahrat historie listu ze kterych se fakturuje
        xSloupecPopisPolozky = 2        'sloupec ve kterem je popis polozek v SPP

        'Protokoly----------------------------------------
        xCisloListu = 1     'radek ve kterem je cislo listu na protokolu
        yCisloListu = 9     'sloupec ve kterem je cislo listu na protokolu
        yStartRadek = 10    'radek od ktereho zacina vycet polozek
        xProtAktivSo = 1    'sloupec ve kterem je oznaceni so na protokolu
        yProtAktivSo = 7    'radek ve kterem je oznaceni so na protokolu
        xCisloSloupcePolozek = 1    'Sloupec ve kterem jsou cisla polozek na protokolu
        xSloupecMnozstviVPoli = 8      'Sloupec ve kterem se helda fakturovane mnozstvi na protokolu
        xSloupecSVyhledat = 2           'Sloupec do ktereho se ma vlozit formule pro svyhledat popisu polozky z SPP
        xSloupecSVyhledatZustatek = 11  'Sloupec do ktereho se ma vlozit formule pro svyhledat zustatkoveho mnozstvi z SPP
        NadpisKontrolaZustatek = "Zustatek dane polozky v SPP:" 'Nadpis k zustatkovemu mnozstvi
        xSloupecMJHlavicka = 9      'sloupec ve kterém je MJ
        xSloupecKonecPopisu = 7     'v jakém sloupci končí popis polozky t těle protokolu od prvního sloupce popisu

        xSloupecOdkazPopis = 2  'slopes odkazu tabuly v SVYHLEDAT s popisem
        xSloupecOdkazMJ = 4  'slopes odkazu tabuly v SVYHLEDAT s mernou jednotkou
        xSloupecOdkazZustatek = 15  'slopes odkazu tabuly v SVYHLEDAT s zustatkem
    End Function
    Function PuleniPopisu()
        i = 11

        Do While i < 200
            If Cells(i, 1) <> "" Then

                Cells(i, 4) = Mid(Cells(i, 4), 1, Len(Cells(i, 4)) / 2)
                Cells(i, 4) = Replace(Cells(i, 4), vbCrLf, "")
            End If
            i = i + 1
        Loop

    End Function
    Function dopln_celk_cislo_protokolu()
        i = 1
        Do While List(i) > Empty
            Sheets(List(i)).Cells(xCisloListu, yCisloListu) = i '& " / " & cislo_listu - 1  prikaz pro doplneni cisla listu do xls
            i = i + 1
        Loop
    End Function
    Function nacti_spp_polozky()
        Sheets(aktiv_list).Select

        NazevListSPP = Sheets(aktiv_list).Name

        PocetObsazenychRadku = Cells(Rows.Count, xSloupecPolozek).End(xlUp).Row
        If Cells(1, 50) <> "opravaProbehla" Then
            OpravaSPP()
        End If

        Erase pokam_zformatovat
        Erase polozka
        i = yZacatekPolozekSPP  '<<<<<<<<<<<<<<<------- 1. --- Uprava pro SPP, odkud zacit hledat polozky -------------
        zacatek_polozek = xSloupecPolozek   '<upravit na kterou polozkou zacina spp

        Do While Cells(i, xSloupecPolozek) <> zacatek_polozek
            i = i + 1
            zacatek_polozek = Cells(i, xSloupecPolozek)
        Loop
        yZacatekPolozekSPP = i

        suma_prazdnych_radku = 0
        cislo_polozky = 0

        'obdobi_fakturace = Cells(i - 2, 19)

        Do While suma_prazdnych_radku < 60

            If Cells(i, xSloupecPolozek) = Empty Then
                suma_prazdnych_radku = suma_prazdnych_radku + 1

            Else
                If IsNumeric(Cells(i, xSloupecPolozek)) = True Then
                    suma_prazdnych_radku = 0
                    cislo_polozky = cislo_polozky + 1
                    polozka(cislo_polozky, 0) = Cells(i, xSloupecPolozek)     'kopirovani cisla nebo nazvu polozky
                    polozka(cislo_polozky, 1) = Cells(i, xSloupecZustatekMnozstvi)  'kopirování množství ze zůstatku
                    polozka(cislo_polozky, 2) = Replace(Cells(i, xSloupecMJ), " ", "")   'kopirovaní MJ
                    polozka(cislo_polozky, 3) = i               'cislo radku na kterem je polozka
                End If
            End If

            i = i + 1
        Loop
        suma_polozek = i - 60

    End Function
    Function projdi_listy()
        celk_pocet_listu = Sheets.Count
        SkutecnyList = 0
        Erase pol_list
        cislo_listu = 1
        StartPol = yStartRadek

        For i = 1 To celk_pocet_listu
            Sheets(i).Select
            If Replace(Cells(yProtAktivSo, xProtAktivSo), " ", "") = aktiv_so Then 'jedná se o list protokolu? /ANO/ -> hledej polozky a scitej
                '// Zjisti kolik je polozek na listu

                Range("A" & StartPol & ":A48").NumberFormat = "General"

                If InStr(1, Sheets(i).Name, "vzor", vbBinaryCompare) > 0 Then
                    GoTo preskoc_list
                End If
                'list neni vzor a zároven je listem protokolu:

                SkutecnyList = SkutecnyList + 1
                List(SkutecnyList) = i  'cislo listo v excelu, na kterém již je legitimní list protokolu

                Erase mem_pol
                Erase znak

                'Napsani cisla polozky do praveho horniho rohu je v samostatne funkci
                'Cells(1, 9) = cislo_listu
                'cislo_listu = cislo_listu + 1

                polozka_v_zahlavi = 0 '/slouzi k pocitani radku zahlavi
                n_polozek = 0       '/n_polozek slouzi pro pocitani mnozstvi polozek v zahlavi
                'InStr(1,"skut",Cells(10 + polozka_v_zahlavi, 1),vbTextCompare
                Do While InStr(1, Cells(StartPol + polozka_v_zahlavi, xCisloSloupcePolozek), "skut", vbTextCompare) = 0
                    If Cells(StartPol + polozka_v_zahlavi, xCisloSloupcePolozek) <> Empty Then
                        'zkontroluje jestli je svyhledat a kdyztak nahraje funkci pro popis polozky
                        VlozSVyhledatHlavicka(StartPol + polozka_v_zahlavi, xCisloSloupcePolozek)

                        'zkontroluje jestli je svyhledat a kdyztak nahraje funkci pro vypsani zustatku
                        VlozSVyhledatZustatek(StartPol + polozka_v_zahlavi, xSloupecSVyhledatZustatek)
                    End If

                    '> doplnit kod pro kontrolu pritomnosti svyhledat
                    polozka_v_zahlavi = polozka_v_zahlavi + 1

                    If Cells(StartPol + polozka_v_zahlavi - 1, xCisloSloupcePolozek) <> Empty Then
                        n_polozek = n_polozek + 1
                        mem_pol(n_polozek, 0) = polozka_v_zahlavi                   'poradi polozky v zahlavi
                        mem_pol(n_polozek, 1) = Cells(yStartRadek - 1 + polozka_v_zahlavi, 1)  'cislo polozky v zahlavi
                        'znak(polozka_v_zahlavi) = "=B" & 12 + polozka_v_zahlavi
                    End If
                    pokam_zformatovat(i) = polozka_v_zahlavi + StartPol - 1

                    If pokam_zformatovat(i) > 200 Then
                        GoTo preskoc_list   'pokud je daný list prázdný preskočí list
                    End If

                Loop
                '// Projdi protokol a zapis pocty k daným položkám
                'VlozSVyhledat   'vlozi svyhledat do radku

                j_pol = 1
                For j = 1 To 38     'prochází prvním sloupcem a hledá znaky položek

                    If j_pol > n_polozek Then
                        Exit For
                    End If

                    cislo_odkazu = Format(Cells(StartPol + j + polozka_v_zahlavi, xCisloSloupcePolozek), "0") '//převedení na řetězec pro porovnání
                    tmp_mem_pol = Format(mem_pol(j_pol, 1), "0")    '//převedení na řetězec pro porovnání

                    If tmp_mem_pol = Empty Then
                        Exit For
                    End If
                    'cislo_odkazu.NumberFormat = "0"
                    'cislo_odkazu = Cells(13 + j + polozka_v_zahlavi, 2)

                    If cislo_odkazu = tmp_mem_pol And Cells(StartPol + j + polozka_v_zahlavi, xCisloSloupcePolozek).Font.Bold = True Then 'pokud najde cislo polozky stejne jako ma u



                        'Cells(12 + j + polozka_v_zahlavi, 2) = znak(j_pol)
                        For k = 1 To 40

                            '// Ukončení for pokud nenajde pod danou polozkou celkové množství (nebo je mnozství=0) -> zabranuje nacteni celk mnozstvi dalsi polozky v poradi

                            'If Cells(StartPol + k + polozka_v_zahlavi, 8) = 0 And Cells(StartPol + j + k + polozka_v_zahlavi, 1) > Empty Then
                            'j_pol = j_pol + 1
                            'Exit For
                            'End If

                            Cells(StartPol + j + k + polozka_v_zahlavi, xSloupecMnozstviVPoli).Select
                            '// je bunka tucna?
                            If Selection.Font.Bold = True Then
                                tucne = 1
                            Else
                                tucne = 0
                            End If
                            '//je bunka dvakrat podtrhnuta?
                            If Selection.Font.Underline = xlUnderlineStyleSingle Then
                                podtrzene = 1
                            Else
                                podtrzene = 0
                            End If
                            '// Jedna se o cislo v bunce?
                            If IsNumeric(Cells(StartPol + j + k + polozka_v_zahlavi, xSloupecMnozstviVPoli)) = True Then

                                If tucne = 1 And podtrzene = 1 Then
                                    If Cells(StartPol + j + k + polozka_v_zahlavi, xSloupecMnozstviVPoli) <> 0 Then '// Je hodnota v bunce vetsi jak 0?
                                        '// v pripade ze je nazev polozky retezec znaku nebo cislo (napr. 7D)

                                        For n = 1 To suma_polozek

                                            tmp_polozka = Format(polozka(n, 0), "0") '//Převedení čísla položky na řetězec pro účel porovnávání

                                            If tmp_mem_pol = "13" Then '//Debug - hledani pricitani polozek
                                                stopka = 15
                                                If Cells(StartPol + j + k + polozka_v_zahlavi, 11) = 62 Then
                                                    stopka = 15
                                                End If
                                            End If

                                            If tmp_polozka = cislo_odkazu Then

                                                'polozka 4 - Fakturovane mnozstvi
                                                polozka(n, 4) = polozka(n, 4) + Cells(StartPol + j + k + polozka_v_zahlavi, xSloupecMnozstviVPoli)
                                                'polozka 5 - Dosavadni seznam fakturovanych listu
                                                polozka(n, 5) = polozka(n, 5) & Sheets(i).Name & ","
                                                polozka(n, 6) = polozka(n, 6) & "+'" & Sheets(i).Name & "'!" &
                                                                Num2Alpha(xSloupecMnozstviVPoli) &
                                                                str(StartPol + j + k + polozka_v_zahlavi)

                                                Exit For
                                            End If

                                        Next n
                                        j_pol = j_pol + 1
                                        Exit For


                                    End If
                                End If
                                If j_pol > polozka_v_zahlavi Then
                                    Exit For
                                End If

                            End If


                        Next k
                    End If
                Next j

                stranka_protokolu = i
                '//Formatovani stranky
                'Kontrola_radku_protokolu
                '//Konec formatovani stranky
            End If

preskoc_list:

        Next i

        CelkovyPocetListu = cislo_listu - 1
    End Function

    Function vypis_soucty_do_spp()

        Sheets(aktiv_list).Select
        pocet_fakturovanych_polozek = 0

        suma_prazdnych_radku = 0
        cislo_polozky = 0
        i = yZacatekPolozekSPP
        n = 0
        Do While suma_prazdnych_radku < 450



            If Cells(i, xSloupecPolozek) = Empty Then

                suma_prazdnych_radku = suma_prazdnych_radku + 1

            Else
                If IsNumeric(Cells(i, xSloupecPolozek)) = True Then
                    suma_prazdnych_radku = 0
                    'suma_prazdnych_radku = suma_prazdnych_radku + 1
                    n = n + 1
                    old_n = n
                    j = n

                    Do While Cells(i, xSloupecPolozek) <> polozka(j, 0)
                        j = j + 1
                    Loop

                    If j > old_n Then
                        n = j
                    End If

                    If Cells(i, xSloupecPolozek) = polozka(n, 0) Then
                        If polozka(n, 4) <> 0 Then
                            suma_prazdnych_radku = 0

                            'Polozka 4 je mnozstvi do sledovaneho obdobi

                            'Cells(i, xSloupecSledovaneObdobi) = polozka(n, 4)

                            Formula = "=" & Replace(Mid(polozka(n, 6), 2, Len(polozka(n, 6))), " ", "")
                            Range(Num2Alpha(xSloupecSledovaneObdobi) & i & ":" & Num2Alpha(xSloupecSledovaneObdobi) & i).Formula = Formula
                            'Polozka 5 je seznam fakturace
                            Cells(i, xSloupecHistorieFakturace) = polozka(n, 5)
                            'Cells(i, 16).Interior.Color = aktiv_barva

                            If polozka(n, 4) <> Empty Then
                                pocet_fakturovanych_polozek = pocet_fakturovanych_polozek + 1
                            End If
                        Else
                            Cells(i, xSloupecSledovaneObdobi) = ""
                            Cells(i, xSloupecHistorieFakturace) = ""
                        End If
                    End If
                End If
            End If
            i = i + 1
        Loop

    End Function
    Function VlozSVyhledatHlavicka(Radek As Integer, Sloupec As Integer)
        Dim Formula As String
        Dim formSTR As String
        'xSloupecOdkazPopis = 3  'slopes odkazu tabuly v SVYHLEDAT s popisem
        'xSloupecOdkazMJ = 4  'slopes odkazu tabuly v SVYHLEDAT s mernou jednotkou
        'xSloupecOdkazZustatek = 16  'slopes odkazu tabuly v SVYHLEDAT s zustatkem

        If Cells(Radek, Sloupec).HasFormula = False Then
            formSTR = "=VLOOKUP(" & Num2Alpha(Sloupec) & Radek & ",'" & NazevListSPP &
            "'!" & Num2Alpha(xSloupecPolozek) & yZacatekPolozekSPP & ":" &
            Num2Alpha(xSloupecPopisPolozky) & PocetObsazenychRadku & "," & xSloupecOdkazPopis & ")"

            Range(Num2Alpha(xSloupecSVyhledat) & Radek & ":" & Num2Alpha(xSloupecKonecPopisu) & Radek).Formula = formSTR
        End If

        If Cells(Radek, xSloupecMJHlavicka).HasFormula = False Then
            formSTR = "=VLOOKUP(" & Num2Alpha(Sloupec) & Radek & ",'" & NazevListSPP &
        "'!" & Num2Alpha(xSloupecPolozek) & yZacatekPolozekSPP & ":" &
        Num2Alpha(xSloupecMJ) & PocetObsazenychRadku & "," & xSloupecOdkazMJ & ")"

            Range(Num2Alpha(xSloupecMJHlavicka) & Radek & ":" & Num2Alpha(xSloupecMJHlavicka) & Radek).Formula = formSTR
        End If

    End Function

    'xSloupecMJHlavicka


    Function VlozSVyhledatZustatek(Radek As Integer, Sloupec As Integer)
        Dim Formula As String
        Dim formSTR As String
        'Kontrola jestli je nadpis zustatku
        If Cells(yStartRadek - 1, xSloupecSVyhledatZustatek) <> NadpisKontrolaZustatek Then
            Cells(yStartRadek - 1, xSloupecSVyhledatZustatek) = NadpisKontrolaZustatek

            With Cells(yStartRadek - 1, xSloupecSVyhledatZustatek)
                .Font.Bold = True
                .Font.Underline = xlUnderlineStyleSingle
            End With
        End If

        If Cells(Radek, Sloupec).HasFormula = False Then
            formSTR = "=VLOOKUP(" & Num2Alpha(xCisloSloupcePolozek) & Radek & ",'" & NazevListSPP &
            "'!" & Num2Alpha(xSloupecPolozek) & yZacatekPolozekSPP & ":" &
            Num2Alpha(24) & PocetObsazenychRadku & "," & xSloupecOdkazZustatek & ")"

            Range(Num2Alpha(Sloupec) & Radek & ":" & Num2Alpha(Sloupec) & Radek).Formula = formSTR
        End If

    End Function
    Function IsFormula(cell_ref As Range)
        IsFormula = cell_ref.HasFormula
    End Function

    Function sirka_sloupcu()
        'Columns(1).ColumnWidth = 7.71
        nazevlistu = ActiveSheet.Name
        umisteni = ThisWorkbook.Path

        If nacteny_radek = 1 Then
            Radek(1) = "13.5"
            Radek(2) = "25.25"
            Radek(3) = "25.25"
            Radek(4) = "8.25"
            Radek(5) = "25.25"
            Radek(6) = "5.25"
            Radek(7) = "20"
            Radek(8) = "20"
            Radek(9) = "9.75"
            Radek(10) = "18.75"
            Radek(11) = "18.75"
            Radek(12) = "18.75"
            Radek(13) = "16.50"
            Radek(14) = "16.50"
            Radek(15) = "16.50"
            Radek(16) = "16.50"
            Radek(17) = "16.50"
            Radek(18) = "16.55"
            Radek(19) = "16.50"
            Radek(20) = "16.50"
            Radek(21) = "16.50"
            Radek(22) = "18.75"
            Radek(23) = "14.25"
            Radek(24) = "14.25"
            Radek(25) = "14.25"
            Radek(26) = "14.25"
            Radek(27) = "14.25"
            Radek(28) = "14.25"
            Radek(29) = "14.25"
            Radek(30) = "14.25"
            Radek(31) = "14.25"
            Radek(32) = "14.25"
            Radek(33) = "14.25"
            Radek(34) = "14.25"
            Radek(35) = "14.25"
            Radek(36) = "14.25"
            Radek(37) = "14.25"
            Radek(38) = "14.25"
            Radek(39) = "14.25"
            Radek(40) = "14.25"
            Radek(41) = "14.25"
            Radek(42) = "14.25"
            Radek(43) = "14.25"
            Radek(44) = "14.25"
            Radek(45) = "14.25"
            Radek(46) = "14.25"
            Radek(47) = "14.25"
            Radek(48) = "14.25"
            Radek(49) = "14.25"
            Radek(50) = "14.25"
            Radek(51) = "14.25"
            Radek(52) = "14.25"
            Radek(53) = "14.25"
            Radek(54) = "14.25"
            Radek(55) = "14.25"
            Radek(56) = "14.25"
            Radek(57) = "14.25"
            Radek(58) = "14.25"
            Radek(59) = "6.0"
            Radek(60) = "14.25"
            Radek(61) = "6.0"
            nacteny_radek = 0
        End If

        If nacteny_sloupec = 1 Then
            Sloupec(1) = "2.33"
            Sloupec(2) = "7.71"
            Sloupec(3) = "7.29"
            Sloupec(4) = "7.14"
            Sloupec(5) = "9.29"
            Sloupec(6) = "7.57"
            Sloupec(7) = "9"
            Sloupec(8) = "7.14"
            Sloupec(9) = "9.71"
            Sloupec(10) = "9.14"
            Sloupec(11) = "14.29"
            Sloupec(12) = "9.29"
            nacteny_sloupec = 0
        End If

        For i = 1 To 61
            Rows(i).RowHeight = Radek(i)
        Next i

        For i = 1 To 12
            Columns(i).ColumnWidth = Sloupec(i)
        Next i

    End Function

    Function Uprava_pro_tisk()

        ActiveWindow.Zoom = 100
        ActiveWindow.View = xlPageBreakPreview
        Range("B2:L61").Select
        ActiveSheet.PageSetup.PrintArea = "$B$2:$L$61"
        ActiveWindow.SmallScroll Down:=-50
ActiveWindow.SmallScroll ToLeft:=50
Range("A1").Select
        On Error GoTo uznastaveno
        ActiveSheet.VPageBreaks(1).DragOff Direction:=xlToRight, RegionIndex:=1
uznastaveno:
        Range("B2:L61").Select
        With ActiveSheet.PageSetup
            .PrintArea = "$B$2:$L$61"
            .FitToPagesWide = 1
            .FitToPagesTall = 1
            '.VPageBreaks.Location = Range("M1")
        End With


    End Function

    Function Kontrola_radku_protokolu()
        If Cells(60, 7) <> obdobi_fakturace Then

            i = 40

            Do While Cells(i, 2) <> "Za zhotovitele:"
                i = i + 1
            Loop
            za_zhotovitele = i - 1

            Do While Cells(i, 2) <> " V Brně"
                i = i + 1
            Loop

            Cells(i, 7) = obdobi_fakturace

            Delta = Abs(60 - i)
            zmena_radku = za_zhotovitele & ":" & za_zhotovitele

            If i <= 60 Then

                Do While Delta > 0
                    Rows(zmena_radku).Insert
                    Delta = Delta - 1
                Loop
            Else
                Do While Delta > 0
                    Rows(zmena_radku).Delete
                    za_zhotovitele = za_zhotovitele - 1
                    zmena_radku = za_zhotovitele & ":" & za_zhotovitele
                    Delta = Delta - 1
                Loop
            End If
        Else

            Cells(60, 7) = obdobi_fakturace

        End If

        'sirka_sloupcu
        'Ohraniceni
        'Uprava_pro_tisk
    End Function
    Function Ohraniceni()
        ohraniceni_horizont_thin("B13:L" & pokam_zformatovat(stranka_protokolu))
        ohraniceni_O("B11:B" & pokam_zformatovat(stranka_protokolu))
        ohraniceni_O("C11:J" & pokam_zformatovat(stranka_protokolu))
        ohraniceni_O("L12:L" & pokam_zformatovat(stranka_protokolu))
        odstran_ohraniceni("A62:N70")

        ohraniceni_O("B2:L61")
    End Function


    Function ohraniceni_O(oblast)
        Range(oblast).Select

        With Selection.Borders(xlEdgeLeft)
            .Weight = xlMedium
        End With
        With Selection.Borders(xlEdgeTop)
            .Weight = xlMedium
        End With
        With Selection.Borders(xlEdgeBottom)
            .Weight = xlMedium
        End With
        With Selection.Borders(xlEdgeRight)
            .Weight = xlMedium
        End With
    End Function
    Function odstran_ohraniceni(oblast)
        Range(oblast).Select

        Selection.Borders(xlInsideVertical).LineStyle = xlNone '//
        Selection.Borders(xlInsideHorizontal).LineStyle = xlNone '//

        Selection.Borders(xlEdgeTop).LineStyle = xlNone
        Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    End Function
    Function ohraniceni_horizont_thin(oblast)
        Range(oblast).Select
        Selection.Borders(xlDiagonalDown).LineStyle = xlNone
        Selection.Borders(xlDiagonalUp).LineStyle = xlNone
        With Selection.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With Selection.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With Selection.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With Selection.Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
    End Function
    Function zapis_zakrouhlit()

        i = 20
        Do While i < 60

            On Error Resume Next

            retezec = Cells(i, 4)
            If InStr(1, retezec, "=", vbBinaryCompare) > 0 Then

                If InStr(1, Cells(i + 1, 4), "=", vbBinaryCompare) > 0 Then

                    zacatek_scitani = i
                    j = 0
                    Do While Cells(i + j, 4) <> Empty

                        retezec = Cells(i + j, 4)
                        retezec = Replace(retezec, "=", "")
                        retezec = Replace(retezec, ",", ".")
                        retezec = "=ROUND(" & retezec & ",3)"

                        Cells(i, 12).NumberFormat = "@"
                        Cells(i, 12).Font.Bold = False
                        Cells(i, 12).Font.Underline = xlUnderlineStyleSingle

                        Cells(i + j, 11).NumberFormat = "0.000"
                        Cells(i + j, 11).Font.Bold = False
                        Cells(i + j, 11).Font.Underline = xlUnderlineStyleSingle
                        Cells(i + j, 11).Formula = retezec
                        j = j + 1
                        If Cells(i + j, 2) <> Empty Then
                            GoTo preskoc_radek
                        End If
                    Loop

preskoc_radek:

                    konec_scitani = j + i - 1
                    'Cells(i + j, 11).Font.Bold = False
                    'Cells(i + j, 11).Font.Underline = xlUnderlineStyleSingle
                    'Cells(i + j, 11).Formula = "=SUM(K" & zacatek_scitani & ":K" & konec_scitani & ")"

                    i = i + j

                Else 'pouze jeden radek s =

                    retezec = Replace(retezec, "=", "")
                    retezec = Replace(retezec, ",", ".")
                    retezec = "=ROUND(" & retezec & ",3)"
                    Cells(i, 11).NumberFormat = "0.000"
                    Cells(i, 11).Font.Bold = True
                    Cells(i, 11).Font.Underline = xlUnderlineStyleSingle

                    Cells(i, 12).NumberFormat = "@"
                    Cells(i, 12).Font.Bold = True
                    Cells(i, 12).Font.Underline = xlUnderlineStyleSingle

                    Cells(i, 11).Formula = retezec

                End If
            End If


            i = i + 1
        Loop
    End Function

    Function poskladej_string(zaklad As String, pridat As String) As String
        If zaklad <> Empty And pridat = Empty Then 'varianta kdy pridavam retezce prvni str
            poskladej_string = zaklad
        End If

        If zaklad <> Empty And pridat <> Empty Then 'varianta kdy k zakladu pridavam dalsi str
            i = 0
            Do While Char <> "," 'retezec hledajici posledni carku od konce
                Char = Mid(zaklad, Len(zaklad) - i, 1)
                i = i + 1
            Loop

            lastnumber = Mid(zaklad, Len(zaklad) - i, i)

        End If
    End Function


    '////////------------------------------------------\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    Function OpravaSPP()
        For i = yZacatekPolozekSPP To PocetObsazenychRadku
            TMP = Cells(i, xSloupecPolozek)
            Cells(i, xSloupecPolozek).Select
            Cells(i, xSloupecPolozek) = TMP
        Next i
        Cells(1, 50) = "opravaProbehla"
    End Function

    Function VlozSVyhledat()
        i = 10
        Do While Cells(i, 1) <> Empty
            If InStr(1, Cells(i, 1), "skut", vbTextCompare) = 0 Then
                Range("B" & i & ":G" & i).Select
                'Cells(i, 2) = "=VLOOKUP(A" & Str(i) & ",'soupis SŽDC'!$A$12:$T$1000,3)"
                ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-1],'soupis SŽDC'!R12C1:R1000C20,3)"
                i = i + 1
            Else
                Exit Do
            End If
        Loop
    End Function

    Function Alpha2Number(str As String) As Integer
        'PURPOSE: Convert a given letter into it's corresponding Numeric Reference
        'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault
        Dim ColumnNumber As Long
        Dim ColumnLetter As String

        'Input Column Letter
        'ColumnLetter = str

        'Convert To Column Number
        'ColumnNumber = Range(ColumnLetter & 1).Column
        Alpha2Number = Range(str & 1).Column

    End Function
    Function Num2Alpha(cislo As Integer) As String
        'PURPOSE: Convert a given number into it's corresponding Letter Reference
        'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault

        Dim ColumnNumber As Long
        Dim ColumnLetter As String

        'Input Column Number
        'ColumnNumber = cislo

        'Convert To Column Letter
        'ColumnLetter = Split(Cells(1, ColumnNumber).Address, "$")(1)

        'Display Result
        Num2Alpha = Split(Cells(1, cislo).Address, "$")(1)

    End Function



End Class
