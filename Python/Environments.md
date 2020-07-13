# Python Virtual Environments

## virtualenv

```shell
usage: virtualenv [--version] [--with-traceback] [-v | -q] [--app-data APP_DATA] [--clear-app-data] [--discovery {builtin}] [-p py] [--creator {builtin,cpython3-posix,venv}] [--seeder {app-data,pip}] [--no-seed]
                  [--activators comma_sep_list] [--clear] [--system-site-packages] [--symlinks | --copies] [--no-download | --download] [--extra-search-dir d [d ...]] [--pip version] [--setuptools version] [--wheel version] [--no-pip]
                  [--no-setuptools] [--no-wheel] [--symlink-app-data] [--prompt prompt] [-h]
                  dest

optional arguments:
  --version                     display the version of the virtualenv package and its location, then exit
  --with-traceback              on failure also display the stacktrace internals of virtualenv (default: False)
  --app-data APP_DATA           a data folder used as cache by the virtualenv (default: /home/prp1277/.local/share/virtualenv)
  --clear-app-data              start with empty app data folder (default: False)
  -h, --help                    show this help message and exit

verbosity:
  verbosity = verbose - quiet, default INFO, mapping => CRITICAL=0, ERROR=1, WARNING=2, INFO=3, DEBUG=4, NOTSET=5

  -v, --verbose                 increase verbosity (default: 2)
  -q, --quiet                   decrease verbosity (default: 0)

discovery:
  discover and provide a target interpreter

  --discovery {builtin}         interpreter discovery method (default: builtin)
  -p py, --python py            target interpreter for which to create a virtual (either absolute path or identifier string) (default: /home/prp1277/Github-Repos/data-science/bin/python3)

creator:
  options for creator builtin

  --creator {builtin,cpython3-posix,venv}
                                create environment via (builtin = cpython3-posix) (default: builtin)
  dest                          directory to create virtualenv at
  --clear                       remove the destination directory if exist before starting (will overwrite files otherwise) (default: False)
  --system-site-packages        give the virtual environment access to the system site-packages dir (default: False)
  --symlinks                    try to use symlinks rather than copies, when symlinks are not the default for the platform (default: True)
  --copies, --always-copy       try to use copies rather than symlinks, even when symlinks are the default for the platform (default: False)

seeder:
  options for seeder app-data

  --seeder {app-data,pip}       seed packages install method (default: app-data)
  --no-seed, --without-pip      do not install seed packages (default: False)
  --no-download, --never-download
                                pass to disable download of the latest pip/setuptools/wheel from PyPI (default: True)
  --download                    pass to enable download of the latest pip/setuptools/wheel from PyPI (default: False)
  --extra-search-dir d [d ...]  a path containing wheels the seeder may also use beside bundled (can be set 1+ times) (default: [])
  --pip version                 pip version to install, bundle for bundled (default: latest)
  --setuptools version          setuptools version to install, bundle for bundled (default: latest)
  --wheel version               wheel version to install, bundle for bundled (default: latest)
  --no-pip                      do not install pip (default: False)
  --no-setuptools               do not install setuptools (default: False)
  --no-wheel                    do not install wheel (default: False)
  --symlink-app-data            symlink the python packages from the app-data folder (requires seed pip>=19.3) (default: False)

activators:
  options for activation scripts

  --activators comma_sep_list   activators to generate - default is all supported (default: bash,cshell,fish,powershell,python,xonsh)
  --prompt prompt               provides an alternative prompt prefix for this environment (default: None)

config file /home/prp1277/.config/virtualenv/virtualenv.ini missing (change via env var VIRTUALENV_CONFIG_FILE)
```

## pipenv

```shell
Usage: pipenv [OPTIONS] COMMAND [ARGS]...

Options:
  --where             Output project home information.
  --venv              Output virtualenv information.
  --py                Output Python interpreter information.
  --envs              Output Environment Variable options.
  --rm                Remove the virtualenv.
  --bare              Minimal output.
  --completion        Output completion (to be evald).
  --man               Display manpage.
  --support           Output diagnostic information for use in GitHub issues.
  --site-packages     Enable site-packages for the virtualenv.  [env var:
                      PIPENV_SITE_PACKAGES]
  --python TEXT       Specify which version of Python virtualenv should use.
  --three / --two     Use Python 3/2 when creating virtualenv.
  --clear             Clears caches (pipenv, pip, and pip-tools).  [env var:
                      PIPENV_CLEAR]
  -v, --verbose       Verbose mode.
  --pypi-mirror TEXT  Specify a PyPI mirror.
  --version           Show the version and exit.
  -h, --help          Show this message and exit.


Usage Examples:
   Create a new project using Python 3.7, specifically:
   $ pipenv --python 3.7

   Remove project virtualenv (inferred from current directory):
   $ pipenv --rm

   Install all dependencies for a project (including dev):
   $ pipenv install --dev

   Create a lockfile containing pre-releases:
   $ pipenv lock --pre

   Show a graph of your installed dependencies:
   $ pipenv graph

   Check your installed dependencies for security vulnerabilities:
   $ pipenv check

   Install a local setup.py into your virtual environment/Pipfile:
   $ pipenv install -e .

   Use a lower-level pip command:
   $ pipenv run pip freeze

Commands:
  check      Checks for security vulnerabilities and against PEP 508 markers
             provided in Pipfile.
  clean      Uninstalls all packages not specified in Pipfile.lock.
  graph      Displays currently-installed dependency graph information.
  install    Installs provided packages and adds them to Pipfile, or (if no
             packages are given), installs all packages from Pipfile.
  lock       Generates Pipfile.lock.
  open       View a given module in your editor.
  run        Spawns a command installed into the virtualenv.
  shell      Spawns a shell within the virtualenv.
  sync       Installs all packages specified in Pipfile.lock.
  uninstall  Un-installs a provided package and removes it from Pipfile.
  update     Runs lock, then sync.
```

venv

`pipenv --venv`
