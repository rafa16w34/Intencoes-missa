import os


def database_context(request):
    """
    Disponibiliza informações sobre o banco de dados ativo em todos os templates.
    Permite exibir avisos visuais claros quando o modo de teste estiver em uso.
    """
    db_name = os.environ.get('DJANGO_DB', 'db.sqlite3')
    is_test_mode = db_name != 'db.sqlite3'
    return {
        'DB_NAME': db_name,
        'IS_TEST_MODE': is_test_mode,
    }
