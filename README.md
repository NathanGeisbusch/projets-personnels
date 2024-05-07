# ⚡️ Automatic Shutdown ⚡️

Programme écrit en [Zig](https://ziglang.org) permettant de déclencher automatiquement l'extinction d'un ordinateur sous Windows après une certaine période d'inactivité.

## Compilation

```bash
zig build-exe main.zig -O ReleaseSmall -mcpu sandybridge -target x86_64-windows -lc
```

## Installation

Pour lancer le programme automatiquement au démarrage:

- Windows + R
- shell:startup
- Placer l'exécutable dans le dossier
