import click
import os
from pptvideo import pptx2video

@click.command()
@click.option("--pptx", help="Specify the pptx file", prompt="pptx file")
@click.option("--lang", help="Specify the language", default="en", prompt="language")
@click.option("--destmp4", help="Specify the output mp4 file", prompt="destmp4")
def p2v(pptx, lang, destmp4):
  pptx2video(pptx, lang, destmp4) 
