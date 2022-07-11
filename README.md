# VBA
E-kharid filling
Sub Kharid2()

    Dim IE As Object
    Dim doc As HTMLDocument
    Set IE = CreateObject("InternetExplorer.Application")
    
    IE.Visible = True
    IE.Navigate "https://ekharid.haryana.gov.in/FCI/CMR_AA_DAction"
    
    Do While IE.Busy
         Application.Wait DateAdd("s", 1, Now)
    Loop
    Set doc = IE.Document
     'IE.doc.getElementById("u_0_3").Value = "test"
    doc.getElementById("ContentPlaceHolder1_grd_Heads_lbl_PassedAmount_0").Value = doc.getElementById("ContentPlaceHolder1_grd_Heads_lbl_ClaimedAmount_0").innerText
    doc.getElementById("ContentPlaceHolder1_grd_Heads_lbl_PassedAmount_1").Value = doc.getElementById("ContentPlaceHolder1_grd_Heads_lbl_ClaimedAmount_1").innerText
    doc.getElementById("ContentPlaceHolder1_grd_Heads_lbl_PassedAmount_2").Value = doc.getElementById("ContentPlaceHolder1_grd_Heads_lbl_ClaimedAmount_2").innerText
    doc.getElementById("ContentPlaceHolder1_grd_Heads_lbl_PassedAmount_3").Value = doc.getElementById("ContentPlaceHolder1_grd_Heads_lbl_ClaimedAmount_3").innerText
    doc.getElementById("ContentPlaceHolder1_grd_Heads_lbl_PassedAmount_4").Value = doc.getElementById("ContentPlaceHolder1_grd_Heads_lbl_ClaimedAmount_4").innerText
    doc.getElementById("ContentPlaceHolder1_grd_Heads_lbl_PassedAmount_5").Value = doc.getElementById("ContentPlaceHolder1_grd_Heads_lbl_ClaimedAmount_5").innerText
    doc.getElementById("ContentPlaceHolder1_grd_Heads_lbl_PassedAmount_6").Value = doc.getElementById("ContentPlaceHolder1_grd_Heads_lbl_ClaimedAmount_6").innerText
    doc.getElementById("ContentPlaceHolder1_grd_Heads_lbl_PassedAmount_7").Value = doc.getElementById("ContentPlaceHolder1_grd_Heads_lbl_ClaimedAmount_7").innerText
    doc.getElementById("ContentPlaceHolder1_grd_Heads_lbl_PassedAmount_8").Value = doc.getElementById("ContentPlaceHolder1_grd_Heads_lbl_ClaimedAmount_8").innerText
    doc.getElementById("txt_Reason").Value = "PASSED"
End Sub



