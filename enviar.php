<?php
use PHPMailer\PHPMailer\PHPMailer;
use PHPMailer\PHPMailer\Exception;

require 'vendor/autoload.php'; // Asegúrate de que este archivo está en tu servidor

if ($_SERVER["REQUEST_METHOD"] == "POST") {
    $nombre = htmlspecialchars($_POST["nombre"]);
    $email = filter_var($_POST["email"], FILTER_SANITIZE_EMAIL);
    $asunto = htmlspecialchars($_POST["asunto"]);
    $mensaje = htmlspecialchars($_POST["mensaje"]);

    $mail = new PHPMailer(true);

    try {
        // Configuración del servidor SMTP de Gmail
        $mail->isSMTP();
        $mail->Host       = 'smtp.gmail.com';
        $mail->SMTPAuth   = true;
        $mail->Username   = 'tuemail@gmail.com'; // Cambia esto a tu correo
        $mail->Password   = 'tupassword_o_app_password'; // Usa una contraseña de aplicación
        $mail->SMTPSecure = PHPMailer::ENCRYPTION_STARTTLS;
        $mail->Port       = 587;

        // Configuración del correo
        $mail->setFrom($email, $nombre);
        $mail->addAddress('tuemail@gmail.com', 'Destinatario'); // Cambia esto al correo donde recibirás los mensajes

        $mail->Subject = $asunto;
        $mail->Body    = "Nombre: $nombre\nCorreo: $email\n\nMensaje:\n$mensaje";

        $mail->send();
        echo "Mensaje enviado correctamente.";
    } catch (Exception $e) {
        echo "Error al enviar el mensaje: {$mail->ErrorInfo}";
    }
} else {
    echo "Método de solicitud no válido.";
}
?>
