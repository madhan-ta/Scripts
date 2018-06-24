# we will use the Add-Type cmdlet and specify the WinForms .NET assembly name.
Add-Type -AssemblyName System.Windows.Forms
# create an instance of form object
$Form = New-Object System.Windows.Forms.Form
# create label object to create a title
$labeltitle = New-Object system.windows.forms.label
$labeltitle.text = 'my first one'
$form.controls.Add($labeltitle)
# show the form
$form.ShowDialog()
