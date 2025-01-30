import streamlit as st

# Nombres originales de los archivos (asegúrate de que estén en la misma carpeta que el script)
imagen_personal = "Imagen de WhatsApp 2024-11-13 a las 23.46.15_58c7ab86.jpg"
video_personal = "Video de WhatsApp 2025-01-30 a las 12.18.39_bc275536.mp4"

# Estilo de fondo con CSS (Color conche de vino + corazones <3)
page_bg_img = """
<style>
    body {
        background-color: #800020;
        color: white;
        font-family: Arial, sans-serif;
    }
    .main {
        background: rgba(255, 255, 255, 0.15);
        padding: 20px;
        border-radius: 15px;
        text-align: center;
        color: white;
    }
    /* Corazones flotando en el fondo */
    body::before {
        content: "<3    <3    <3    <3    <3    <3    <3";
        font-size: 30px;
        font-weight: bold;
        color: pink;
        position: fixed;
        top: 10%;
        left: 5%;
        white-space: nowrap;
        opacity: 0.5;
    }
    body::after {
        content: "<3    <3    <3    <3    <3    <3    <3";
        font-size: 30px;
        font-weight: bold;
        color: pink;
        position: fixed;
        bottom: 10%;
        right: 5%;
        white-space: nowrap;
        opacity: 0.5;
    }
</style>
"""

st.markdown(page_bg_img, unsafe_allow_html=True)

# Inicializar estado de la conversación y el tamaño del botón
if "estado" not in st.session_state:
    st.session_state.estado = "inicio"
if "tamanio_si" not in st.session_state:
    st.session_state.tamanio_si = 20  # Tamaño inicial del botón "Sí"
if "intentos_no" not in st.session_state:
    st.session_state.intentos_no = 0  # Contador de rechazos

# Contenedor con diseño bonito
st.markdown("<div class='main'>", unsafe_allow_html=True)

st.title("💖 ¿Would You Be My Valentine? 💖")

# ✅ Mostrar la imagen directamente (sin `try-except` para evitar bloqueos)
st.image(imagen_personal, caption="De nuestro viajecito a Quito gg love u amor 💕", use_container_width=True)

# Lógica para manejar respuestas
if st.session_state.estado == "inicio":
    st.subheader("Special Invitation ❤️")
    
    col1, col2 = st.columns([1, 1])
    with col1:
        if st.button("Yes, si quiero y siempre yes! 💘", key="si_1", help="Haz clic para aceptar!", use_container_width=True):
            st.session_state.estado = "aceptado"
    with col2:
        if st.button("No... 😢", key="no_1", use_container_width=True):
            st.session_state.estado = "seguro"
            st.session_state.intentos_no += 1  # Incrementar contador de rechazos

elif st.session_state.estado == "seguro":
    mensajes_no = [
        "¿Segura/o? Podemos ir a michael´s :3... 🥺",
        "Piensa en todas las flores que te daría... 🌹💌",
        "¿En serio me vas a romper el corazón así? 💔😭",
        "Vamos, dime que sí y te prometo una sorpresa especial... 🎁💖",
        "¿Ni aunque te invite a una cena romántica? 🍽️🥂",
        "F... 💕✨",
        "Di que sí, ILY 😍🔥"
    ]
    from random import choice
    mensaje_aleatorio = choice(mensajes_no)
    
    st.subheader(mensaje_aleatorio)
    
    # Incrementar el tamaño del botón "Sí" cada vez que dice "No"
    st.session_state.tamanio_si += 10  

    col1, col2 = st.columns([1, 1])
    with col1:
        if st.button(f"Sí, quiero! 💘", key=f"si_{st.session_state.intentos_no}", 
                     help="No puedes resistirte 😘", use_container_width=True):
            st.session_state.estado = "aceptado"
    with col2:
        if st.button("No... 😭", key=f"no_{st.session_state.intentos_no}", use_container_width=True):
            st.session_state.estado = "seguro"
            st.session_state.intentos_no += 1  # Suma otro intento

elif st.session_state.estado == "aceptado":
    st.balloons()
    st.subheader("💖 ¡Yujuuu! Sabía que dirías que sí 🥰💖")
    st.write("¡Nos espera un San Valentín increíble juntos! 🌹✨")
    
    # ✅ Agregar el video directamente sin `try-except`
    st.video(video_personal)

st.markdown("</div>", unsafe_allow_html=True)
