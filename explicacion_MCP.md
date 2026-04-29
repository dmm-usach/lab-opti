# El Problema del Clique Máximo (MCP)

## ¿Qué es un Clique?

Imagina una red social donde cada persona es un **nodo** y cada amistad es una **conexión**. Un **clique** es un grupo de personas donde **todos se conocen entre sí** — no hay nadie en el grupo que no sea amigo de todos los demás.

El **Problema del Clique Máximo** pregunta: ¿cuál es el grupo más grande posible con esta propiedad?

---

## Una analogía concreta

Supón que tienes 5 personas: Ana, Bruno, Carla, Diego y Elena.

```
Ana ── Bruno ── Carla
 │       │
Diego ── Elena
```

- Ana, Bruno y Diego se conocen todos entre sí → eso es un **clique de tamaño 3**
- Si agregas a Elena, pero Elena no conoce a Ana → ya **no es clique**

El objetivo es encontrar el grupo más numeroso donde todos se conocen.

---

## ¿Por qué es difícil?

El MCP es un problema **NP-duro**, lo que significa que:

- No existe ningún algoritmo conocido que lo resuelva **rápidamente** para grafos grandes
- A medida que el grafo crece, el tiempo de cómputo crece **exponencialmente**
- Ni siquiera es posible *aproximar* la solución a un factor constante en tiempo polinomial, a menos que P = NP

En términos prácticos: un grafo de 500 nodos ya puede tardar **horas** en resolverse de forma exacta.

---

