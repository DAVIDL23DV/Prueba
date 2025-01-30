import streamlit as st
import random

# Imagen de fondo con CSS
page_bg_img = """
<style>
    body {
        background-image: url("https://images.unsplash.com/photo-1516641394681-8c35c322b412");
        background-size: cover;
        background-position: center;
        background-attachment: fixed;
    }
    .main {
        background: rgba(255, 255, 255, 0.8);
        padding: 20px;
        border-radius: 15px;
        text-align: center;
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

# Lista de mensajes aleatorios para cuando diga "No"
mensajes_no = [
    "¿Segura/o? Yo tenía un chocolate para ti... 🍫🥺",
    "Piensa en todas las flores y cartas que te daría... 🌹💌",
    "¿En serio me vas a romper el corazón así? 💔😭",
    "Vamos, dime que sí y te prometo una sorpresa especial... 🎁💖",
    "¿Ni aunque te invite a una cena romántica? 🍽️🥂",
    "No lo pienses demasiado, el amor está en el aire... 💕✨",
    "Di que sí, ¡somos el match perfecto! 😍🔥"
]

# Contenedor con diseño bonito
st.markdown("<div class='main'>", unsafe_allow_html=True)

st.title("💖 ¿Quieres ser mi San Valentín? 💖")

# Lógica para manejar respuestas con imágenes y cambios en el tamaño del botón "Sí"
if st.session_state.estado == "inicio":
    st.image("https://i.pinimg.com/originals/5b/15/57/5b155775e580a3039bb0f3e9acbc154e.gif", width=300)  # Imagen inicial
    st.subheader("Esta es una invitación especial para ti ❤️")
    
    col1, col2 = st.columns([1, 1])
    with col1:
        if st.button("Sí, quiero! 💘", key="si_1", help="Haz clic para aceptar!", use_container_width=True):
            st.session_state.estado = "aceptado"
    with col2:
        if st.button("No... 😢", key="no_1", use_container_width=True):
            st.session_state.estado = "seguro"
            st.session_state.intentos_no += 1  # Incrementar contador de rechazos

elif st.session_state.estado == "seguro":
    st.image("https://media1.tenor.com/m/TPbczMykUzIAAAAC/crying.gif", width=300)  # Imagen triste
    mensaje_aleatorio = random.choice(mensajes_no)  # Elegir un mensaje aleatorio
    st.subheader(mensaje_aleatorio)
    
    # Incrementar el tamaño del botón "Sí" cada vez que dice "No"
    st.session_state.tamanio_si += 10  

    col1, col2 = st.columns([1, 1])
    with col1:
        if st.button(f"Sí, quiero! 💘", key=f"si_{st.session_state.intentos_no}", 
                     help="No puedes resistirte 😘", use_container_width=True, 
                     args=(st.session_state.tamanio_si,)):
            st.session_state.estado = "aceptado"
    with col2:
        if st.button("No... 😭", key=f"no_{st.session_state.intentos_no}", use_container_width=True):
            st.session_state.estado = "seguro"
            st.session_state.intentos_no += 1  # Suma otro intento

elif st.session_state.estado == "aceptado":
    st.balloons()
    st.image("https://media1.tenor.com/m/cFDbD6jZxDoAAAAd/cute-love.gif", width=300)  # Imagen feliz
    st.subheader("💖 ¡Yujuuu! Sabía que dirías que sí 🥰💖")
    st.write("¡Nos espera un San Valentín increíble juntos! 🌹✨")

st.markdown("</div>", unsafe_allow_html=True)
