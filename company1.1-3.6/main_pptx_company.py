"""Entry point for generating the company PPT."""
from content_pptx_company import PPTContentEngine
from full_pptx_company import PPTFullEngine


def main():
    print("Generating company PPT...")
    engine = PPTFullEngine(PPTContentEngine())
    engine.generate()
    print("Done.")


if __name__ == "__main__":
    main()

