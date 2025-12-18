"""Entry point for generating the governance PPT demo (A/B layouts only)."""
from config_pptx import PPT_CONFIG
from content_pptx import PPTContentEngine
from full_pptx import PPTFullEngine


def main():
    print("Generating governance PPT with A/B templates...")
    engine = PPTFullEngine(PPTContentEngine())
    engine.generate()
    print("Done.")


if __name__ == "__main__":
    main()

