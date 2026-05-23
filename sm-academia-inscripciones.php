<?php
/**
 * Plugin Name: SM Academia — Inscripciones a Google Sheets
 * Description: Envía las inscripciones del formulario Jetpack al GAS de SM Academia.
 * Version:     1.0.0
 * Author:      SM Academia
 *
 * INSTALACIÓN:
 *   1. Subir este archivo a /wp-content/plugins/sm-academia-inscripciones/
 *      (crear la carpeta si no existe)
 *   2. Activar el plugin en WP Admin → Plugins
 *
 * REQUISITO:
 *   El GAS debe tener la acción NUEVA_INSCRIPCION en doPost.
 *   Ver instrucciones en el chat de SM Academia.
 */

if ( ! defined( 'ABSPATH' ) ) exit;

// ── URL del GAS ────────────────────────────────────────────────────────────────
define( 'SM_GAS_URL', 'https://script.google.com/macros/s/AKfycbwZOhEUs3dbIGGWLEdx5gMMFD7L85B5yyIdxwRqdZkZcikYNwC7cH3bPiGGudOfBko7zQ/exec' );

// ── Email del administrador para notificaciones de error ───────────────────────
define( 'SM_ADMIN_EMAIL', 'asmarriv1986@gmail.com' );

/**
 * Se dispara cuando Jetpack Contact Form envía un formulario.
 *
 * @param int   $post_id    ID de la página que contiene el formulario.
 * @param array $mail       Configuración del correo del formulario.
 * @param array $all_values Array asociativo [ 'Etiqueta del campo' => 'valor' ].
 */
add_action( 'grunion_after_message_sent', 'sm_inscripcion_a_gas', 10, 3 );
function sm_inscripcion_a_gas( $post_id, $mail, $all_values ) {

    // Solo actuar en la página de inscripción
    if ( get_post_field( 'post_name', $post_id ) !== 'inscripcion-academia-sm-futbol' ) {
        return;
    }

    // Días preferidos puede venir como array (checkboxes) o string
    $dias_raw = $all_values['Días preferidos'] ?? '';
    if ( is_array( $dias_raw ) ) {
        $dias = implode( ', ', array_map( 'sanitize_text_field', $dias_raw ) );
    } else {
        $dias = sanitize_text_field( $dias_raw );
    }

    $data = [
        'action'             => 'NUEVA_INSCRIPCION',
        'nombre_alumno'      => sanitize_text_field( $all_values['Nombre del jugador/a']       ?? '' ),
        'apellidos_alumno'   => sanitize_text_field( $all_values['Apellidos del jugador/a']    ?? '' ),
        'fecha_nac'          => sanitize_text_field( $all_values['Fecha nacimiento (AAAA-MM-DD)'] ?? '' ),
        'colegio'            => sanitize_text_field( $all_values['Colegio']                    ?? '' ),
        'observaciones'      => sanitize_textarea_field( $all_values['Observaciones']          ?? '' ),
        'nombre_tutor'       => sanitize_text_field( $all_values['Nombre del tutor/a']         ?? '' ),
        'apellidos_tutor'    => sanitize_text_field( $all_values['Apellidos del tutor/a']      ?? '' ),
        'telefono'           => sanitize_text_field( $all_values['Teléfono']                   ?? '' ),
        'email'              => sanitize_email(      $all_values['Correo electrónico']         ?? '' ),
        'servicio'           => sanitize_text_field( $all_values['Servicio']                   ?? '' ),
        'modalidad'          => sanitize_text_field( $all_values['Modalidad']                  ?? '' ),
        'preferencia_centro' => sanitize_text_field( $all_values['Preferencias']               ?? '' ),
        'dias'               => $dias,
        'autoriza_imagen'    => sanitize_text_field( $all_values['Autorizo uso de imagen']     ?? 'Sí' ),
        'mensaje'            => sanitize_textarea_field( $all_values['Mensaje']                ?? '' ),
    ];

    // Llamada al GAS — no bloqueante: el usuario no espera respuesta
    $response = wp_remote_post( SM_GAS_URL, [
        'headers'  => [ 'Content-Type' => 'application/json; charset=utf-8' ],
        'body'     => wp_json_encode( $data ),
        'timeout'  => 30,
        'blocking' => false,   // respuesta ignorada; fallo silencioso en frontend
    ] );

    // Log de error (solo visible en debug.log si WP_DEBUG_LOG está activo)
    if ( is_wp_error( $response ) ) {
        error_log( '[SM Academia] Error al llamar al GAS: ' . $response->get_error_message() );
    }
}
