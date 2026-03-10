from comissoes.database.models import registrar_email_envio, listar_historico_email
registrar_email_envio(2, "rep000571@example.com", "erro", "unitario")
print(listar_historico_email())
