shebang := if os() == 'windows' { 'pwsh.exe' } else { '/usr/bin/env pwsh' }
set shell := ["nu", "-c"]
set windows-shell := ["pwsh.exe", "-NoLogo", "-NoProfile","-Command"]
set dotenv-load := true
set script-interpreter := ["pwsh.exe", "-NoLogo", "-NoProfile","-Command"]
set dotenv-filename	:= ".env"
set unstable
set fallback
# set dotenv-required := true
export JUST_ENV := "just_env" # WARN: this is also a method to export env var. 
_default:
    @just --list --unsorted
    
alias b := build
[group('dev')]
build:
    # build task here

alias r := run
default_args := 'args here'
[group('debug')]
run args=default_args:
    @Write-Host {{default_args}} -ForegroundColor Red

alias fmt := format
format:
    # format plesase. could also run rfmt

alias t := test
[group('dev')]
test:
    # test.

alias w := watch
watch:
    # watch, mostly spin up an application e.g. `r r`-> read log by less/bat/ov/tailspin

alias dep := deploy
deploy:
    # deploy ...like at least commit and push to remote first.

alias sk := seek
seek:
    # seek ...what? maybe reference to other?

[script]
script:
    Write-Host "this is in powershell, without shebang syntax"
