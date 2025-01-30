import streamlit as st

# Establecer el fondo con una imagen de corazones (usando CSS)
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
    }
</style>
"""

st.markdown(page_bg_img, unsafe_allow_html=True)

# Inicializar estado de la conversación
if "estado" not in st.session_state:
    st.session_state.estado = "inicio"

# Contenedor con diseño bonito
st.markdown("<div class='main'>", unsafe_allow_html=True)

st.title("💖 ¿Quieres ser mi San Valentín? 💖")

# Lógica para manejar respuestas
if st.session_state.estado == "inicio":
    st.subheader("Esta es una invitación especial para ti ❤️")
    if st.button("Sí, quiero! 💘"):
        st.session_state.estado = "aceptado"
    if st.button("No... 😢"):
        st.session_state.estado = "seguro"

elif st.session_state.estado == "seguro":
    st.subheader("¿Estás segura/o? 🥺💔")
    if st.button("Sí, quiero! 💘"):
        st.session_state.estado = "aceptado"
    if st.button("No... 😭"):
        st.session_state.estado = "seguro"

elif st.session_state.estado == "aceptado":
    st.balloons()
    st.subheader("💖 ¡Yujuuu! Sabía que dirías que sí 🥰💖")
    st.write("¡Nos espera un San Valentín increíble juntos! 🌹✨")

st.markdown("</div>", unsafe_allow_html=True)
