Function Get-FileName($initialDirectory)
{
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory

    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.filename
}



$global:inputfile="empty"
$global:inputfile2="empty2"





   function ConnectMySQL([string]$user,[string]$pass,[string]$MySQLHost,[string]$database) {
 
# Load MySQL .NET Connector Objects
[void][system.reflection.Assembly]::LoadWithPartialName("MySql.Data")
 
# Open Connection
$connStr = "server=" + $MySQLHost + ";port=3306;uid=" + $user + ";pwd=" + $pass + ";database="+$database+";Pooling=FALSE"
$conn = New-Object MySql.Data.MySqlClient.MySqlConnection($connStr)
$conn.Open()
$cmd = New-Object MySql.Data.MySqlClient.MySqlCommand("USE $database", $conn)
return $conn
 
}
 
 
 
 $Global:excel="" 
 $Global:wb="" 
 $Global:sh="" 
 


 #tworzymy instancjê okna
$form = New-Object System.Windows.Forms.Form
#okno otrzymuje nazwê, rozmiar oraz pozycjê (œrodek ekranu)
#podobne w³aœciwoœci otrzymuj¹ pozosta³e obiekty wrzucane na okno
$licznik=0
$form.Text = "Excel to mySql converter"
$form.Size = New-Object System.Drawing.Size(470,300)
$form.StartPosition = "CenterScreen"


$button1 = New-Object System.Windows.Forms.Button
$button1.Text="Wybierz plik do zapisu"
$button1.Size = New-Object System.Drawing.Size(100,50)
$button1.Location = New-Object System.Drawing.Size(100,100)
$form.Controls.Add($button1)

 

$button3 = New-Object System.Windows.Forms.Button
$button3.Text="Zapisz dane w bazie danych"
$button3.Size = New-Object System.Drawing.Size(100,50)
$button3.Location = New-Object System.Drawing.Size(250,100)
$form.Controls.Add($button3)

$Label1 = New-Object System.Windows.Forms.Label
$Label1.Text =  $global:inputfile
$Label1.AutoSize = $True
$Label1.Location = New-Object System.Drawing.Size(180,200)
$Form.Controls.Add($Label1)

$Label2 = New-Object System.Windows.Forms.Label
$Label2.Text =  "Wybrany plik: "
$Label2.AutoSize = $True
$Label2.Location = New-Object System.Drawing.Size(180,185)
$Form.Controls.Add($Label2)



$image = [System.Drawing.Image]::Fromfile('C:\\temp\\dane1.jpg')    
$pictureBox = new-object Windows.Forms.PictureBox  
$pictureBox.width=470
$pictureBox.height=300
$pictureBox.top=0
$pictureBox.left=0
$pictureBox.Image=$image

 $form.Controls.add($pictureBox)                  
 
 $Form.Add_Shown({$Form.Activate()})
 

  

    

[System.Collections.Arraylist] $Global:arrayList=@()
[System.Collections.Arraylist] $Global:a=@()

Function loadXmlFile($file){
   
 write-host ">>"$file
 

 
 for($i=0;$i -le $file.Length;$i++){
 

 if($file[$i] -eq "."){
 break;}
 $xml_file+=$file[$i]

 }

$xml_file+=".xml"





 #[xml]$TablesAttribute = Get-Content   C:\temp\info.xml
[xml]$TablesAttribute = Get-Content  $xml_file  


 $war1 = $TablesAttribute.Tables.Table 
 $war1.ChildNodes.count
 
 for($i=0;$i -le $war1.ChildNodes.Count;$i++){
    $arrayList.Add(  $war1.ChildNodes.Item($i))
       }
 
 $ilosc = $war1.ChildNodes.count 

 New-Item c:\temp\tempInfo.txt -ItemType file
  
 Clear-Content "c:\temp\tempInfo.txt" 
 
    $arrayList >> "c:\temp\tempInfo.txt" 
 
 Write-host "Ilosc: "$ilosc
 $Global:a=(Get-Content "c:\temp\tempInfo.txt") | Select-Object -last  $ilosc

 Write-host "zamienna a: " $a[0]
                        
  Remove-Item c:\temp\tempInfo.txt                           

}

function WriteMySQLQuery($conn, [string]$query) {
 
$command = $conn.CreateCommand()
$command.CommandText = $query
$RowsInserted = $command.ExecuteNonQuery()
$command.Dispose()
if ($RowsInserted) {
return $RowInserted
} else {
return $false
}
}


Function getCellFromExcel($w, $k)
{ 

 

    $kom=$Global:sh.Cells.Item($w,$k).Text


return $kom

}


Function insertExcelData(){
                          # setup vars
                        $user = 'root'
                        $pass = ''
                        $database = 'syop'
                        $MySQLHost = '127.0.0.1'
 
                        # Connect to MySQL Database
                        $conn = ConnectMySQL $user $pass $MySQLHost $database
                        # Read all the records from table
$numberOfColumns=($a.Count-5)/2
Write-Host "kolumn: " $numberOfColumns 

 for($k=0; $k -lt $a[3]-$a[1] +1 ; $k++){



$kumulNames=""
$kumulValues=""
 for($i=0; $i -lt $numberOfColumns*2  ; $i=$i+2){

if($i -eq $numberOfColumns*2-2){
$kumulNames+=$a[5+$i] + "  "
}
else{
$kumulNames+=$a[5+$i] + " , "
}

 }
 
 for($j=0; $j -lt $numberOfColumns  ; $j=$j+1){
 
$row=0
 $row=$k+$a[1] 
 #Write-Host "row:>" $row "<"

 
 $col=0
# $col=$j+1
$col=$j+$a[2]

 if($j -eq $numberOfColumns-1){
$kumulValues+= " "" "+(getCellFromExcel $row $col)+" "" "
}
else{
$kumulValues+= " "" "+(getCellFromExcel $row $col)+" "" , "
}

 }

  

$insertQuery ="INSERT INTO "  +$a[0] + " ( "+$kumulNames+" ) VALUES "+ " ( " +$kumulValues+ " ) "

 
  #Write-Host "Test Query: " $insertQuery 


                         $Rows = WriteMySQLQuery $conn $insertQuery
                         Write-Host "Executed: " $insertQuery

}
$Global:excel.Workbooks.Close()
Stop-Process -processname  Excel

}




Function createTable(){




$tableName=$a[0]
 
        

 $query="CREATE TABLE ``syop``."+$tableName
 
$kumul=""
     


          for($i=5 ; $i -lt $a.Count-1; $i=$i+2){

            if($i-eq $a.count-2){

                $kumul+=$a[$i] +" "+$a[$i+1] +"NULL" }

            else {
                $kumul+=$a[$i] +" "+$a[$i+1]+   "NULL , "
            }

            
            }


$query=$query+"("+ $kumul+")"

 

   
    # zapytanie tworzace tebelê
      

                        # setup vars
                        $user = 'root'
                        $pass = ''
                        $database = 'syop'
                        $MySQLHost = '127.0.0.1'
 
                        # Connect to MySQL Database
                        $conn = ConnectMySQL $user $pass $MySQLHost $database
                        # Read all the records from table
                       
                       
                     $query   
                      
 
                        $Rows = WriteMySQLQuery $conn $query 
                        Write-Host "Executed for create a table: " $query

}

 

                        $button1.add_Click(
                        {
                          $global:inputfile = Get-FileName "c:/"
                          Write-Host "Excel: "$inputfile
                          $Label1.Text =   $inputfile



                         $Global:excel=new-object -com excel.application
                        $Global:wb=$Global:excel.workbooks.open($inputfile)
                        $Global:sh=$Global:wb.Sheets.Item(1)
    

                            loadXmlFile  $inputfile
 
                        })

 

                      

  $button3.add_Click({
#TO musi byœ zawsze
 createTable 
   insertExcelData
    #  insertData                     
                        })

 

$form.ShowDialog();
######################################################################################################