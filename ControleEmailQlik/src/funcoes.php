<?php
 
$opcao = isset($_POST['opcao']) ? $_POST['opcao'] : '';

$dir = 'C:\xampp\htdocs\ControleEmailQlik\src\VBScriptQlikSend\acessos';

if (!empty($opcao)){   
    switch ($opcao)
    {
        case 'folder':
            {
                echo getAllFiles();
                break;
            }
     	case 'report':
        	{
        		$file = isset($_POST['file']) ? $_POST['file'] : '';
            echo getAllReports($file);
            break;
        	}
        case 'emailapp':
      		{

      			$file = isset($_POST['file']) ? $_POST['file'] : '';
      			echo getAllEmailApp($file);
      			break;
      		}
      	case 'emailreport':
      		{
      			$file = isset($_POST['file']) ? $_POST['file'] : '';
      			$report = isset($_POST['report']) ? $_POST['report'] : '';

      			if(empty($report)){
      				echo getAllEmailApp($file);
      			}
      			else{
      				echo getAllEmailReport($file,$report);
      			}
      			
      			break;
      		}
      	 case 'modifyEmail':
      		{
      			$emailsend = isset($_POST['emailsend']) ? $_POST['emailsend'] : '';
      			$file = isset($_POST['file']) ? $_POST['file'] : '';
      			$report = isset($_POST['report']) ? $_POST['report'] : '';
      			echo modifyEmail($emailsend,$file,$report);
      			break;
      		}
          case 'reporttable':
          {

            $reporttable = isset($_POST['reporttable']) ? $_POST['reporttable'] : '';
            echo getAllFilesReport();
            break;
          }
    }
}

function getAllFiles(){

	   global $dir;
       $cdir = scandir($dir); 

       foreach ($cdir as $key => $value) 
       { 
       	 if (!in_array($value,array(".",".."))) 
          { 
             if (is_dir($dir . DIRECTORY_SEPARATOR . $value)) 
             { 

                if(substr($value,-4) == '-pst'){
                 
                  echo "<option value='" . $value . "'>".ucwords(mb_strtolower(substr($value,0, strpos($value, '-pst'))))."</option>"; 
                
                }
                
             } 
              
          }
       }     
  }


function getAllReports($file){
    
	   global $dir;
 
     $dirinternal = ($dir."/".$file);
	   $cdir = scandir($dirinternal); 

       foreach ($cdir as $key => $value)  { 
       	
          if (!in_array($value,array(".",".."))) 
          { 
             if (is_dir($dirinternal . DIRECTORY_SEPARATOR . $value)) 
             { 
                $result[$value] = dirToArray($dirinternal . DIRECTORY_SEPARATOR . $value); 
             } 
             else 
             {  
					    echo "<option value='" . $value . "'>".ucwords(mb_strtolower(substr($value,0, strpos($value, '.txt'))))."</option>"; 
					
                
             } 
          }
       } 
 }

function getAllEmailApp($file){

     global $dir;
 
     $dirinternal = ($dir."/".$file);
     $cdir = scandir($dirinternal); 
     $stremails = '';

       foreach ($cdir as $key => $value)  { 
        
          if (!in_array($value,array(".",".."))) 
          { 
             if (is_dir($dirinternal . DIRECTORY_SEPARATOR . $value)) 
             { 
                $result[$value] = dirToArray($dirinternal . DIRECTORY_SEPARATOR . $value); 
             } 
             else 
             {    
                    $dirinternalfile = ($dirinternal."/".$value);
                    $fn = fopen($dirinternalfile,"r");
                    $stremails = $stremails . fgets($fn);
             } 
          }
       }

     $stremails = str_replace("\n", "", $stremails);
     $emails = explode(";", $stremails);
     $emails = array_unique($emails);
    
     foreach ($emails as &$email) {
        if(!empty($email)){
       echo "<option value='" . $email . "'>".$email."</option>"; 
     }
    }
 }

 function getAllEmailReport($file,$report){


     global $dir;
 
     $dirinternalfile = ($dir."/".$file."/".$report);
     
     $fn = fopen($dirinternalfile,"r");
     $stremails = $stremails . fgets($fn);

     $stremails = str_replace("\n", "", $stremails);
     $emails = explode(";", $stremails);
     $emails = array_unique($emails);
    
     foreach ($emails as &$email) {
       if(!empty($email)){
        echo "<option value='" . $email . "'>".$email."</option>"; 
       }
     }
 }

function modifyEmail($emailsend,$file,$report){

     global $dir;
 
     $dirinternalfile = ($dir."/".$file."/".$report);
     
     $fp = fopen($dirinternalfile, 'w');
     fwrite($fp, $emailsend);
     fclose($fp);
}


function popularTabela(){
    
     global $dir;

     $cdir = scandir($dir); 

       foreach ($cdir as $key => $valuepaste) 
       { 
         if (!in_array($valuepaste,array(".",".."))) 
          { 
             if (is_dir($dir . DIRECTORY_SEPARATOR . $valuepaste)) 
             { 

                if(substr($valuepaste,-4) == '-pst'){
                 
                       $dirinternal = ($dir."/".$valuepaste);
                       $cdir = scandir($dirinternal); 

                         foreach ($cdir as $key => $valuereport)  { 
                          
                            if (!in_array($valuereport,array(".",".."))) 
                            { 
                               if (is_dir($dirinternal . DIRECTORY_SEPARATOR . $valuereport)) 
                               { 
                                  $result[$valuereport] = dirToArray($dirinternal . DIRECTORY_SEPARATOR . $valuereport); 
                               } 
                               else 
                               {  

                                       $dirinternalfile = ($dir."/".$valuepaste."/".$valuereport);
     
                                       $fn = fopen($dirinternalfile,"r");
                                       $stremails="";
                                       $stremails = $stremails . fgets($fn);

                                       $stremails = str_replace("\n", "", $stremails);
                                       $emails = explode(";", $stremails);
                                       $emails = array_unique($emails);
                                      
                                       foreach ($emails as &$email) {
                                         if(!empty($email)){
                                          echo "<tr>
                                                  <td>".ucwords(mb_strtolower(substr($valuepaste,0, strpos($valuepaste, '-pst'))))."</td>
                                                  <td>".ucwords(mb_strtolower(substr($valuereport,0, strpos($valuereport, '.txt'))))."</td>
                                                  <td>".$email."</td>
                                                </tr>"; 
                                         }
                                       }
    
                               } 
                            }
                         } 
                }
                
             } 
              
          }
       }  
  }

?>