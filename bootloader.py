import compilar_exames

if __name__ == "__main__":
    try:
        compilar_exames.main()
    except Exception as e:
        print("❌ Erro:", e)
        input("Pressione Enter para sair...")
