# How to start up mkdocs

## main requirements
```
pip install mkdocs
pip install mkdocs-material
pip install mike
```

## set up new documentation
```
mkdocs new .
```

## update mkdocs.yml
```
site_name: My Docs

nav:
    homepage: index.md

theme:
  name: material
  custom_dir: overrides #added to overide element e.g. header, outdated, footer

extra:
  version:
    provider: mike
```

## How to use mike

### step 1. create an overides folder and `overrides/main.html`
```
mkdir overrides
cd overrides
touch main.html
```

### step 2. add the following to the `overrides/main.html`
```
{% extends "base.html" %}

{% block outdated %}
  You're not viewing the latest version.
  <a href="{{ '../' ~ base_url }}"> 
    <strong>Click here to go to latest.</strong>
  </a>
{% endblock %}
```

### step 3. build documents with [mike](https://github.com/jimporter/mike)

mike will build your docs in a new branch `gh-pages` each time you use the `mike deploy <version>` command

Do this the first time
```
mike deploy 0.1 latest #build 0.1 version of documentation and alias as latest
mike set-default latest #set the default latest version of documentation to be 'latest'
mike serve #serve the files using mike
mike build 0.1 --push #defaults to origin remote and pushes there
```

building on another version
```
mike deploy 0.2 --update latest #build the documentation in version 0.2 and update alias to latest
mike serve #check your build serving site with mike
mike deploy -0.2 --push #when you are ready deploy and push
```

## Material extensions
[pymarkdown reference](https://facelessuser.github.io/pymdown-extensions/extensions/snippets/)

[mkdocs material extensions](https://squidfunk.github.io/mkdocs-material/setup/extensions/)

[mkdocs admonitions (dynamic expanding sections)](https://squidfunk.github.io/mkdocs-material/reference/admonitions/#supported-types)

### Snippets for bringing in other markdowns into your documentation

adding the snippets extension to mkdocs.yml
```
markdown_extensions:
  - pymdownx.snippets:
      base_path: docs
```

adding to the index.md
```
# inside docs/index.md

--8<-- "../README.md"
```