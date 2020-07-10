<?php
$Username=$_POST['Username']
$PhoneNo=$_POST['Phone No']
$Email=$_POST['Email Id']
$Message=$_POST['Message']
if(!empty($Username)||!empty(PhoneNo)||!empty(Email Id)||!empty(Message))
{
 $host="localhost";
 $dbUsername='root';
 $dbPassword=''
 $dbname='registar'


 $conn=new mysqli($host,$dbUsername,$dbPassword,$dbname);
 if(mysqli_connect_error()){

    die('Connect Error('.mysqli_connect_error().')'.mysqli_connect_error());

 }
 else{
     $SELECT="SELECT Email Id from register where Email Id=?Limit 1";
     $INSERT="INSERT into register(Username,PhoneNo,Email Id,Message) values(?,?,?,?)";

     $stmt=$conn->prepare($SELECT);
     $stmt->$bind_param("s",$Email)
     $stmt->execute();
     $stmt->bind_result($Email)
     $stmt->store_result();
     $rnum=$stmt->num_rows;
      if($rnum==0){

        $stmt->close();
         $stmt=$conn->prepare($INSERT);
         $stmt->$bind_param("siss",$Username,$PhoneNo,$Email);
         $stmt->execute();
         echo"New Record inserted Successfully";
      }
      else{
          echo"someone already registered using";
      }
      $stmt->close();
      $conn->close();
 }
}
else{
    echo"All field are required";
    die();
}


?>