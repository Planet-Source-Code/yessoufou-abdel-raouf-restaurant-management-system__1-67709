Attribute VB_Name = "MdlPullDownMenu"
Public Sub Initialize()
   ' Call frm_Main.ctrl_SkinableForm.LoadSkin(frm_Main)
        
    Call frmScreen.ctrl_PullDownMenu.AddItem("File")
    Call frmScreen.ctrl_PullDownMenu.AddItem("Transactions")
    Call frmScreen.ctrl_PullDownMenu.AddItem("Report")
    Call frmScreen.ctrl_PullDownMenu.AddItem("Administration")
    Call frmScreen.ctrl_PullDownMenu.AddItem("Database")
    Call frmScreen.ctrl_PullDownMenu.AddItem("Tools")
    Call frmScreen.ctrl_PullDownMenu.AddItem("Help")

End Sub
