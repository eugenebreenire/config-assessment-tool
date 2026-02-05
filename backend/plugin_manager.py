import click
import os
import sys
import subprocess

PLUGIN_DIR = "plugins"

@click.group()
def cli():
    pass

@cli.command(name="list")
def list_plugins():
    """List available plugins"""
    if not os.path.exists(PLUGIN_DIR):
        print("No plugins directory found.")
        return

    plugins = [d for d in os.listdir(PLUGIN_DIR) if os.path.isdir(os.path.join(PLUGIN_DIR, d)) and not d.startswith('__')]
    if not plugins:
        print("No plugins found.")
    else:
        print("Available plugins:")
        for p in plugins:
            plugin_path = os.path.join(PLUGIN_DIR, p)
            main_file = os.path.join(plugin_path, "main.py")
            status = "(Ready)" if os.path.exists(main_file) else "(Missing main.py)"
            print(f"  - {p} {status}")

@cli.command()
@click.argument('name')
@click.argument('args', nargs=-1)
def start(name, args):
    """Start a plugin"""
    plugin_path = os.path.join(PLUGIN_DIR, name)
    if not os.path.exists(plugin_path):
        print(f"Plugin '{name}' not found.")
        sys.exit(1)

    main_file = os.path.join(plugin_path, "main.py")
    if not os.path.exists(main_file):
        print(f"Plugin '{name}' does not have a main.py entry point.")
        sys.exit(1)

    # Revert to local logic for pipenv/virtualenv handling or just usage of current python
    # We will assume plugins manage dependencies via requirements.txt if they are standalone
    # and we do a simple check here similar to before.

    python_executable = sys.executable
    requirements_file = os.path.join(plugin_path, "requirements.txt")

    if os.path.exists(requirements_file):
        print(f"Dependency file found: {requirements_file}")
        venv_dir = os.path.join(plugin_path, ".venv")
        # Handle Windows vs Unix venv paths
        if sys.platform == "win32":
            venv_python = os.path.join(venv_dir, "Scripts", "python.exe")
        else:
            venv_python = os.path.join(venv_dir, "bin", "python")

        if not os.path.exists(venv_dir):
            print(f"Creating isolated environment for {name}...")
            subprocess.check_call([sys.executable, "-m", "venv", venv_dir])

            print(f"Installing dependencies from requirements.txt...")
            subprocess.check_call([venv_python, "-m", "pip", "install", "-r", requirements_file])
            print("Dependencies installed.")

        python_executable = venv_python

    print(f"Starting plugin {name}...", flush=True)
    cmd = [python_executable, main_file] + list(args)

    # We must preserve environment variables
    env = os.environ.copy()
    # Add plugin dir to PYTHONPATH so it can import its own modules
    # And current PYTHONPATH
    current_pythonpath = env.get("PYTHONPATH", "")
    env["PYTHONPATH"] = f"{plugin_path}:{current_pythonpath}"

    try:
        subprocess.run(cmd, env=env)
    except KeyboardInterrupt:
        print("\nPlugin execution interrupted.")

if __name__ == "__main__":
    cli()
