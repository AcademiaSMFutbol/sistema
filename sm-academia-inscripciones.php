<?php
/**
 * Plugin Name: SM Academia — Inscripciones a Google Sheets
 * Description: Envía las inscripciones y preinscripciones de los formularios Jetpack al GAS de SM Academia.
 * Version:     1.1.0
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
define( 'SM_GAS_URL', 'https://script.google.com/macros/s/AKfycbxr3x2Ops-a1LlIV8KcSPYFuJMF56ijy8It82Qek6mIHxASdIlJZWIuFwufWE7-YrsBhg/exec' );

// ── Email del administrador para notificaciones de error ───────────────────────
define( 'SM_ADMIN_EMAIL', 'asmarriv1986@gmail.com' );

/**
 * Se dispara cuando Jetpack Contact Form envía un formulario.
 *
 * @param int   $post_id    ID de la página que contiene el formulario.
 * @param array $mail       Configuración del correo del formulario.
 * @param array $all_values Array asociativo [ 'Etiqueta del campo' => 'valor' ].
 */
add_action( 'grunion_after_message_sent', 'sm_formulario_a_gas', 10, 3 );
function sm_formulario_a_gas( $post_id, $mail, $all_values ) {
    $slug = get_post_field( 'post_name', $post_id );

    if ( $slug === 'inscripcion-academia-sm-futbol' ) {
        sm_inscripcion_a_gas( $post_id, $all_values );
    } elseif ( $slug === 'preinscripcion-escaleritas' ) {
        sm_preinscripcion_a_gas( $post_id, $all_values );
    }
}

// ── Inscripción SM Academia Fútbol ──────────────────────────────────────────
function sm_inscripcion_a_gas( $post_id, $all_values ) {

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
        error_log( '[SM Academia] Error al llamar al GAS (inscripción): ' . $response->get_error_message() );
    }
}

// ── Preinscripción de interés — SM Extraescolares Escaleritas ───────────────
function sm_preinscripcion_a_gas( $post_id, $all_values ) {

    // Actividades puede llegar como array (checkboxes múltiples)
    $actividades_raw = $all_values['Actividades de interés'] ?? '';
    if ( is_array( $actividades_raw ) ) {
        $actividades = implode( ', ', array_map( 'sanitize_text_field', $actividades_raw ) );
    } else {
        $actividades = sanitize_text_field( $actividades_raw );
    }

    $dias_raw = $all_values['Días disponibles'] ?? '';
    if ( is_array( $dias_raw ) ) {
        $dias = implode( ', ', array_map( 'sanitize_text_field', $dias_raw ) );
    } else {
        $dias = sanitize_text_field( $dias_raw );
    }

    $data = [
        'action'              => 'NUEVA_PREINSCRIPCION',
        'centro'              => 'ESCALERITAS',
        'curso_academico'     => '2026-2027',
        // Alumno
        'nombre_alumno'       => sanitize_text_field( $all_values['Nombre del alumno/a']         ?? '' ),
        'apellidos_alumno'    => sanitize_text_field( $all_values['Apellidos del alumno/a']      ?? '' ),
        'curso_alumno'        => sanitize_text_field( $all_values['Curso escolar']               ?? '' ),
        'fecha_nac'           => sanitize_text_field( $all_values['Fecha de nacimiento']         ?? '' ),
        'salud'               => sanitize_textarea_field( $all_values['Enfermedades/alergias']   ?? '' ),
        'observaciones'       => sanitize_textarea_field( $all_values['Observaciones']           ?? '' ),
        // Tutor 1
        'nombre_tutor1'       => sanitize_text_field( $all_values['Nombre tutor 1']              ?? '' ),
        'apellidos_tutor1'    => sanitize_text_field( $all_values['Apellidos tutor 1']           ?? '' ),
        'relacion_tutor1'     => sanitize_text_field( $all_values['Relación tutor 1']            ?? '' ),
        'nif_tutor1'          => sanitize_text_field( $all_values['NIF/DNI tutor 1']             ?? '' ),
        'telefono_tutor1'     => sanitize_text_field( $all_values['Teléfono tutor 1']            ?? '' ),
        'email_tutor1'        => sanitize_email(      $all_values['Email tutor 1']               ?? '' ),
        // Tutor 2
        'nombre_tutor2'       => sanitize_text_field( $all_values['Nombre tutor 2']              ?? '' ),
        'apellidos_tutor2'    => sanitize_text_field( $all_values['Apellidos tutor 2']           ?? '' ),
        'relacion_tutor2'     => sanitize_text_field( $all_values['Relación tutor 2']            ?? '' ),
        'nif_tutor2'          => sanitize_text_field( $all_values['NIF/DNI tutor 2']             ?? '' ),
        'telefono_tutor2'     => sanitize_text_field( $all_values['Teléfono tutor 2']            ?? '' ),
        'email_tutor2'        => sanitize_email(      $all_values['Email tutor 2']               ?? '' ),
        // Consentimientos
        'autoriza_imagen'     => sanitize_text_field( $all_values['Cesión de imágenes']          ?? '' ),
        'acepta_datos'        => sanitize_text_field( $all_values['Política de datos']           ?? '' ),
        'acepta_datos_salud'  => sanitize_text_field( $all_values['Consentimiento datos salud']  ?? '' ),
        // AMPA y pago
        'socio_ampa'          => sanitize_text_field( $all_values['¿Desea hacerse socio del AMPA?'] ?? '' ),
        'iban'                => sanitize_text_field( $all_values['IBAN para domiciliación']     ?? '' ),
        // Actividades
        'actividades'         => $actividades,
        'dias_disponibles'    => $dias,
    ];

    $response = wp_remote_post( SM_GAS_URL, [
        'headers'  => [ 'Content-Type' => 'application/json; charset=utf-8' ],
        'body'     => wp_json_encode( $data ),
        'timeout'  => 30,
        'blocking' => false,
    ] );

    if ( is_wp_error( $response ) ) {
        error_log( '[SM Academia] Error al llamar al GAS (preinscripción): ' . $response->get_error_message() );
    }
}
