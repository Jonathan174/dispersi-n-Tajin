import pygame
import sys
import subprocess


# Inicializar Pygame
pygame.init()

# Definir colores
blanco = (255, 255, 255)
negro = (0, 0, 0)
azul = (0, 0, 255)
verde = (108, 153, 35)
rojo = (187, 78, 78)

# Configurar la ventana
ancho, alto = 900, 500
ventana = pygame.display.set_mode((ancho, alto))
pygame.display.set_caption("A Solutions")

# Configurar el ícono de la ventana
icono = pygame.image.load("images\Logo imagen.jpeg")
pygame.display.set_icon(icono)

# Configurar fuentes
fuente_titulo = pygame.font.SysFont("Georgia", 65)
fuente_botones = pygame.font.Font(None, 24)

# Cargar y redimensionar la imagen del logo
imagen_logo = pygame.image.load("images\logotajín.png")
nuevo_ancho = 500  # Ancho deseado
nuevo_alto = 250    # Alto deseado
imagen_logo = pygame.transform.scale(imagen_logo, (nuevo_ancho, nuevo_alto))

#Título
titulo_adicional = fuente_titulo.render("Soluciones para tus archivos", True, blanco)
ventana.blit(titulo_adicional, (70, 200))

# Coordenadas y tamaño del botón
boton_x, boton_y = (ancho - ventana.get_width())//2, (alto - ventana.get_height())//2
boton_ancho, boton_alto = 210, 50

# Bucle principal
while True:
    # Limpiar la pantalla
    ventana.fill(rojo)

    # Dibujar título con imagen
    ventana.blit(imagen_logo, ((ancho - imagen_logo.get_width())//2, 30))

    # Dibujar título "Cálculo de cuotas"
    titulo_calculo_cuotas = fuente_titulo.render("Cálculo de incentivos", True, blanco)
    ventana.blit(titulo_calculo_cuotas, ((ancho - titulo_calculo_cuotas.get_width())//2, alto - imagen_logo.get_height() +50))

    # Definir funciones para dibujar botones
    texto_boton1 = fuente_botones.render("Seleccionar archivo", True, blanco)
    boton_x = (ancho - boton_ancho)//2
    boton_y = alto - (texto_boton1.get_height()*2 + boton_alto)
    pygame.draw.rect(ventana, verde, ((ancho - boton_ancho)//2, alto - (texto_boton1.get_height()*2 + boton_alto), boton_ancho, boton_alto), border_radius=10)
    ventana.blit(texto_boton1, ((ancho - texto_boton1.get_width())//2, alto - (texto_boton1.get_height() + boton_alto)))

    for evento in pygame.event.get():
        if evento.type == pygame.QUIT:
            pygame.quit()
            sys.exit()

        if evento.type == pygame.MOUSEBUTTONDOWN:
            x, y = evento.pos
            if boton_x <= x <= boton_x + boton_ancho and boton_y <= y <= boton_y + boton_alto:
                # Cierra Pygame correctamente
                pygame.quit()       
                subprocess.run(["python", "app.py"])

    # Actualizar la pantalla
    pygame.display.flip()

    # Controlar la velocidad del bucle
    pygame.time.Clock().tick(60)