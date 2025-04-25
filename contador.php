<?php
$archivo = 'visitas.txt';

// Si el archivo no existe, lo crea con valor 0
if (!file_exists($archivo)) {
    file_put_contents($archivo, 0);
}

// Lee el valor actual
$visitas = (int)file_get_contents($archivo);

// Incrementa el contador
$visitas++;

// Guarda el nuevo valor
file_put_contents($archivo, $visitas);

// Devuelve el valor al navegador
echo $visitas;
?>
