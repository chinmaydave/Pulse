from . import create_app


app = create_app()


if __name__ == "__main__":
    cfg = app.config["PULSE_CONFIG"]
    app.run(host=cfg.host, port=cfg.port, debug=True)
