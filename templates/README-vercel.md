# Porque é que esta pasta tem um `vercel.json`

O repositório está ligado a **dois** projetos no Vercel, e ambos fazem deploy
a cada push:

| Projeto        | Endereço                     | Configuração                     |
|----------------|------------------------------|----------------------------------|
| `euromilhoes`  | `euromilhoes-chi.vercel.app` | framework Flask, raiz do repo    |
| `euromilhoes1` | `euromilhoes1.vercel.app`    | `rootDirectory=templates`, estático |

O segundo servia esta pasta como ficheiros soltos: sem Flask, sem `/api/...`
e sem `/guia`, por isso tudo o que não fosse a página inicial dava 404 — e a
própria página chegava com `{{ version }}` por substituir.

O `vercel.json` desta pasta é lido **só** pelo `euromilhoes1`, cuja raiz é
aqui. O projeto principal usa o `/vercel.json` da raiz do repositório e ignora
este. O Flask também o ignora: desta pasta só lhe interessa o `index.html`.

## Notas

- O redirecionamento é **307 e não 301**. Um redirecionamento permanente fica
  em cache nos browsers e custa a desfazer. Para reverter, basta apagar o
  `vercel.json` desta pasta.
- O schema do `vercel.json` **não aceita chaves extra**: uma tentativa de lá
  meter um campo `_comentario` fez o build falhar com
  `should NOT have additional property`. Daí esta explicação viver num
  ficheiro à parte.
- Há ainda uma segunda linha de defesa no `index.html`: a página deteta que
  foi servida crua — o marcador do Jinja chega literal — e redireciona
  sozinha. Fica útil se este `vercel.json` for removido.
