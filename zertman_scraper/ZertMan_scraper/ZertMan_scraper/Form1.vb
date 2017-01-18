Public Class mainForm

    Dim caster As Integer
    Dim startIndex As Integer
    Dim endIndex As Integer
    Dim current As Integer
    Dim urlPrefix As String


    Dim createScript As String = "create table #company" & vbCrLf &
                        "(" & vbCrLf &
                                 "GUID  varchar(50)," & vbCrLf &
                                 "companyName  nvarchar(max)," & vbCrLf &
                                 "mark nvarchar(max)," & vbCrLf &
                                 "workingHours  nvarchar(max)," & vbCrLf &
                                 "workingFields  nvarchar(max)," & vbCrLf &
                                 "contactInfo  nvarchar(max)," & vbCrLf &
                                 "mail  nvarchar(max)," & vbCrLf &
                                 "webPage  nvarchar(max)," & vbCrLf &
                                 "profileFields  nvarchar(max)," & vbCrLf &
                                 "workingArea  nvarchar(max)," & vbCrLf &
                                 "nationalClassification  nvarchar(max)," & vbCrLf &
                                 "subOffices  nvarchar(max)," & vbCrLf &
                                 "serviceCenter  nvarchar(max)," & vbCrLf &
                                 "workers  nvarchar(max)," & vbCrLf &
                                 "workerSexGender  nvarchar(max)," & vbCrLf &
                                 "parentCompany  nvarchar(max)," & vbCrLf &
                                 "childCompany  nvarchar(max)," & vbCrLf &
                                 "creators  nvarchar(max)," & vbCrLf &
                                 "import  nvarchar(max)," & vbCrLf &
                                 "export  nvarchar(max)," & vbCrLf &
                                 "diapazoni  nvarchar(max)," & vbCrLf &
                                 "socialMedia  nvarchar(max)" & vbCrLf &
                         ")" & vbCrLf

    Dim companyInsertTemplate As String = "insert into #company" & vbCrLf &
                                        "select " & vbCrLf &
                                            "newid() ," & vbCrLf &
                                            "N'@companyName@'," & vbCrLf &
                                            "N'@mark@'," & vbCrLf &
                                            "N'@workingHours@'," & vbCrLf &
                                            "N'@workingFields@'," & vbCrLf &
                                            "N'@contactInfo@'," & vbCrLf &
                                            "N'@mail@'," & vbCrLf &
                                            "N'@webPage@'," & vbCrLf &
                                            "N'@profileFields@'," & vbCrLf &
                                            "N'@workingArea@'," & vbCrLf &
                                            "N'@nationalClassification@'," & vbCrLf &
                                            "N'@subOffices@'," & vbCrLf &
                                            "N'@serviceCenter@'," & vbCrLf &
                                            "N'@workers@'," & vbCrLf &
                                            "N'@workerSexGender@'," & vbCrLf &
                                            "N'@parentCompany@'," & vbCrLf &
                                            "N'@childCompany@'," & vbCrLf &
                                            "N'@creators@'," & vbCrLf &
                                            "N'@import@'," & vbCrLf &
                                            "N'@export@'," & vbCrLf &
                                            "N'@diapazoni@'," & vbCrLf &
                                            "N'@socialMedia@'" & vbCrLf &
                                        ""
    Dim mark, workingHours, workingFields, contactInfo, mail, webPage, socialMedia, profileFields, workingArea, nationalClassification, subOffices, serviceCenter, workers, workerSexGender, parentCompany, childCompany, creators, import, export, diapazoni As String
    Dim openFile As System.IO.StreamWriter
    Dim loadCounter As Integer = 0


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        urlPrefix = "http:\\bcat.ge\front\Profile.aspx?ixCompany="


        startIndex = 1
        endIndex = 25000
        current = startIndex + 1


        WebBrowser1.Width = 900
        WebBrowser1.Height = 600


        Dim file As System.IO.StreamWriter

        file = New System.IO.StreamWriter("d:\\zertManScraper.txt", False)
        file.WriteLine("--___________________________________!!! Scraped Script FOR MSSQL!!!____________________________________________")
        file.WriteLine(createScript)
        file.Close()

        openFile = New System.IO.StreamWriter("d:\\zertManScraper.txt", True)

        WebBrowser1.Navigate(urlPrefix & startIndex.ToString)







    End Sub





    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub WebBrowser1_DocumentCompleted(sender As Object, e As WebBrowserDocumentCompletedEventArgs) Handles WebBrowser1.DocumentCompleted
        '
        Dim browser As WebBrowser
        browser = TryCast(sender, WebBrowser)


        RemoveHandler browser.DocumentCompleted, AddressOf WebBrowser1_DocumentCompleted
        MessageBox.Show("You can start Scraping")




    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        If Not Trim(startValue.Text) = "" Then
            startIndex = CInt(startValue.Text)
            current = startIndex + 1
        End If

        If Not Trim(endValue.Text) = "" Then
            endIndex = CInt(endValue.Text)
        End If


        If startIndex < endIndex Then
            Dim wb As WebBrowser = New WebBrowser
            wb.ScriptErrorsSuppressed = True
            AddHandler wb.DocumentCompleted, AddressOf completed
            wb.Navigate(urlPrefix & startIndex.ToString)

        End If






    End Sub

    Sub completed(sender As Object, e As WebBrowserDocumentCompletedEventArgs)

        If loadCounter = 10 Then
            loadCounter = 0
            If startIndex < endIndex Then
                startIndex = startIndex + 1
                Dim navigated As WebBrowser = TryCast(sender, WebBrowser)
                If Not navigated.Document Is Nothing Then

                    Dim contact As HtmlElement = navigated.Document.GetElementById("profile_contact_info_box")

                    Dim mailText As HtmlElement = navigated.Document.GetElementById("profile_contact_email")

                    Dim webSite As HtmlElement = navigated.Document.GetElementById("profile_contact_website")

                    Dim Links As HtmlElement = navigated.Document.GetElementById("profile_social_links")


                    Dim profile As HtmlElement = navigated.Document.GetElementById("profile_public_left_box")
                    Dim employees As HtmlElement = navigated.Document.GetElementById("company_employees")
                    Dim marketing As HtmlElement = navigated.Document.GetElementById("company_marketing_fields")

                    Dim companyProfileFields As HtmlElement = navigated.Document.GetElementById("company_profile_fields")
                    Dim workingAreaGenericHTML As HtmlElement = navigated.Document.GetElementById("company_services")
                    Dim nationalClassificationHTML As HtmlElement = navigated.Document.GetElementById("company_classificators")
                    Dim subOfficesHTML As HtmlElement = navigated.Document.GetElementById("company_branches")
                    Dim serviceCenterHTML As HtmlElement = navigated.Document.GetElementById("company_servicecenters")

                    Dim workersHTML As HtmlElement = navigated.Document.GetElementById("company_employees")
                    Dim workerSexGenderHTML As HtmlElement = navigated.Document.GetElementById("company_GenderAndAge")
                    Dim parentCompanyHTML As HtmlElement = navigated.Document.GetElementById("company_ParentCompany")
                    Dim childCompanyHTML As HtmlElement = navigated.Document.GetElementById("company_ChildCompanies")
                    Dim creatorsHTML As HtmlElement = navigated.Document.GetElementById("company_shareholders")

                    Dim importHTML As HtmlElement = navigated.Document.GetElementById("company_import")

                    Dim exportHTML As HtmlElement = navigated.Document.GetElementById("company_export")

                    Dim diapazoniHTML As HtmlElement = navigated.Document.GetElementById("company_financialFields")

                    ''   PROFILE LEFT   
                    ''Console.WriteLine(profile.InnerText)
                    Dim profileString As String()

                    profileString = profile.InnerText.ToString.Split(vbCrLf)
                    For i = 0 To profileString.Count - 1
                        profileString(i) = profileString(i).Replace(vbLf, "")
                    Next
                    For i = 0 To profileString.Count - 1
                        If profileString(i) = "სავაჭრო მარკა" Then
                            If i + 1 < profileString.Count Then
                                mark = profileString(i + 1)
                            End If
                        End If
                        If profileString(i) = "სამუშაო საათები" Then
                            If i + 1 < profileString.Count Then
                                workingHours = profileString(i + 1)
                            End If
                        End If
                        If profileString(i) = "საქმიანობა" Then
                            If i + 1 < profileString.Count Then
                                workingFields = profileString(i + 1)
                            End If
                        End If
                    Next

                    ''  CONTACT RIGHT

                    Dim b As String()
                    ''b = contact.InnerText.ToString.Split(vbCrLf)
                    contactInfo = contact.InnerText.ToString.Replace("საკონტაქტო ინფორმაცია", "")
                    contactInfo = contactInfo.Trim()

                    ''    MAIL   
                    mail = mailText.InnerText.ToString.Replace("ელ-ფოსტა: ", "")
                    '' WebSite
                    webPage = webSite.InnerText.ToString.Replace("ვებ-გვერდი:", "")

                    ''  Social Media
                    socialMedia = ""
                    Dim linksCollection As HtmlElementCollection = Links.GetElementsByTagName("a")
                    For i = 0 To linksCollection.Count - 1
                        socialMedia = socialMedia & linksCollection(i).GetAttribute("href") & "  ;  "
                    Next

                    ''   Company Profile Fields 

                    profileFields = companyProfileFields.InnerText

                    ''  WorkingArea
                    workingArea = workingAreaGenericHTML.InnerText

                    ''  Classification
                    nationalClassification = nationalClassificationHTML.InnerText

                    ''  sub offices
                    subOffices = subOfficesHTML.InnerText

                    ''  Service Center
                    serviceCenter = serviceCenterHTML.InnerText
                    ''  Menegment

                    workers = workersHTML.InnerText
                    workerSexGender = workerSexGenderHTML.InnerText
                    parentCompany = parentCompanyHTML.InnerText
                    childCompany = childCompanyHTML.InnerText
                    creators = creatorsHTML.InnerText

                    import = importHTML.InnerText
                    export = exportHTML.InnerText
                    diapazoni = diapazoniHTML.InnerText


                    Console.WriteLine("next step ")



                    checkForNulls()
                    printFile()
                    emptyVars()

                    If startIndex < endIndex Then
                        Try
                            navigated.Navigate(urlPrefix & startIndex.ToString)
                        Catch ex As Exception
                            Console.WriteLine(ex.ToString)
                            MessageBox.Show("ERROR at " & startIndex.ToString)

                            openFile.Close()
                        End Try

                    Else
                        openFile.Close()
                        MessageBox.Show("Completed at " & startIndex.ToString)
                    End If


                End If
            End If
        Else
            loadCounter = loadCounter + 1
        End If

    End Sub


    Sub printFile()
        Dim toPrint As String = ""
        toPrint = companyInsertTemplate.Replace("@mark@", mark)
        toPrint = toPrint.Replace("@workingHours@", workingHours)
        toPrint = toPrint.Replace("@workingFields@", workingFields)
        toPrint = toPrint.Replace("@contactInfo@", contactInfo)
        toPrint = toPrint.Replace("@mail@", mail)
        toPrint = toPrint.Replace("@webPage@", webPage)
        toPrint = toPrint.Replace("@socialMedia@", socialMedia)
        toPrint = toPrint.Replace("@profileFields@", profileFields)
        toPrint = toPrint.Replace("@workingArea@", workingArea)
        toPrint = toPrint.Replace("@nationalClassification@", nationalClassification)
        toPrint = toPrint.Replace("@subOffices@", subOffices)
        toPrint = toPrint.Replace("@serviceCenter@", serviceCenter)
        toPrint = toPrint.Replace("@workers@", workers)
        toPrint = toPrint.Replace("@workerSexGender@", workerSexGender)
        toPrint = toPrint.Replace("@parentCompany@", parentCompany)
        toPrint = toPrint.Replace("@childCompany@", childCompany)
        toPrint = toPrint.Replace("@creators@", creators)
        toPrint = toPrint.Replace("@import@", import)
        toPrint = toPrint.Replace("@export@", export)
        toPrint = toPrint.Replace("@diapazoni@", diapazoni)



        openFile.WriteLine("--_________________________________________________________________________")
        openFile.WriteLine(toPrint)




    End Sub

    Sub emptyVars()
        mark=""
        workingHours = ""
        workingFields = ""
        contactInfo = ""
        mail = ""
        webPage = ""
        socialMedia = ""
        profileFields = ""
        workingArea = ""
        nationalClassification = ""
        subOffices = ""
        serviceCenter = ""
        workers = ""
        workerSexGender = ""
        parentCompany = ""
        childCompany = ""
        creators = ""
        import = ""
        export = ""
        diapazoni = ""
    End Sub

    Function checkIfNull(varName)
        If varName = Nothing Then
            Return ""
        Else
            Return varName

        End If
    End Function

    Sub checkForNulls()
        mark = checkIfNull(mark)
        workingHours = checkIfNull(workingHours)
        workingFields = checkIfNull(workingFields)
        contactInfo = checkIfNull(contactInfo)
        mail = checkIfNull(mail)
        webPage = checkIfNull(webPage)
        socialMedia = checkIfNull(socialMedia)
        profileFields = checkIfNull(profileFields)
        workingArea = checkIfNull(workingArea)
        nationalClassification = checkIfNull(nationalClassification)
        subOffices = checkIfNull(subOffices)
        serviceCenter = checkIfNull(serviceCenter)
        workers = checkIfNull(workers)
        workerSexGender = checkIfNull(workerSexGender)
        parentCompany = checkIfNull(parentCompany)
        childCompany = checkIfNull(childCompany)
        creators = checkIfNull(creators)
        import = checkIfNull(import)
        export = checkIfNull(export)
        diapazoni = checkIfNull(diapazoni)
    End Sub

End Class
