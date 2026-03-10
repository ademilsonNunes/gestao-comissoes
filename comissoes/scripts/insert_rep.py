from comissoes.database.models import upsert_representante, listar_representantes
upsert_representante("000571","Rep 000571","rep000571@example.com","Segue sua comissão.")
print(listar_representantes())
