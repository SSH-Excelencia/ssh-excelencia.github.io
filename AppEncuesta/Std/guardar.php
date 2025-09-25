<?php
// Ruta donde guardar
$carpeta = __DIR__ . "data/";

if (!file_exists($carpeta)) {
    mkdir($carpeta, 0777, true);
}

// Recibir datos JSON
$data = json_decode(file_get_contents("php://input"), true);

if ($data && isset($data["nombre"]) && isset($data["contenido"])) {
    $nombreArchivo = basename($data["nombre"]);
    $rutaArchivo = $carpeta . $nombreArchivo;

    // Guardar (append para acumular respuestas en un solo archivo)
    file_put_contents($rutaArchivo, $data["contenido"], FILE_APPEND);
    http_response_code(200);
    echo "Guardado en servidor";
} else {
    http_response_code(400);
    echo "Datos inválidos";
}
?>

