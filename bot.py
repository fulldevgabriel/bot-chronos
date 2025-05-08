import discord
from discord.ext import commands
from datetime import datetime
import os
from dotenv import load_dotenv

from utils import (
    criar_tabela_horarios,
    registrar_entrada,
    registrar_saida,
    criar_resumo_excel
)

load_dotenv()

intents = discord.Intents.default()
intents.message_content = True
bot = commands.Bot(command_prefix='!', intents=intents)

@bot.command()
async def entrada(ctx):
    usuario = str(ctx.author)
    data = datetime.now().strftime('%Y-%m-%d')
    hora = datetime.now().strftime('%H:%M:%S')
    registrar_entrada(usuario, data, hora)
    await ctx.send(f'Entrada registrada para {usuario} às {hora}')

@bot.command()
async def saida(ctx, *, resumo):
    usuario = str(ctx.author)
    data = datetime.now().strftime('%Y-%m-%d')
    hora = datetime.now().strftime('%H:%M:%S')
    registrar_saida(usuario, data, hora, resumo)
    await ctx.send(f'Saída registrada para {usuario} às {hora} com resumo: {resumo}')

@bot.command()
async def criar_resumo(ctx):
    try:
        nome_arquivo = criar_resumo_excel()
        await ctx.send("Aqui está o resumo:", file=discord.File(nome_arquivo))
    except Exception as e:
        await ctx.send(f"Erro ao gerar o Excel: {e}")

criar_tabela_horarios()
bot.run(os.getenv('DISCORD_TOKEN'))
