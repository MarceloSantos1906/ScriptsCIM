import logging
from telegram import (
    Update,
    Message,
)
from telegram.ext import (
    filters,
    MessageHandler,
    ApplicationBuilder,
    CommandHandler,
    ContextTypes,
    CallbackContext,
)
from sgc import main_sgc
import datetime
import pytz
import time
import httpx


class app:
    def __init__(self, token) -> None:
        reply = input("chave é <KEY>?\n")
        if reply[:1].lower() == "s":
            self.chave = "<KEY>"
        else:
            self.chave = input("digite a chave:\n")
        self.senha = input("digite a senha:\n")
        reply = input("impressora é pdpgo125?\n")
        if reply[:1].lower() == "s":
            self.impressora = "<PRINTER>"
        else:
            self.impressora = input("digite a impressora:\n")
        application = (
            ApplicationBuilder().token(token).connection_pool_size(500).build()
        )
        self.now = datetime.datetime.now(pytz.timezone("UTC"))
        self.ulima_verificacao = ""
        self.intervalo = 0
        self.monitorar = False
        self.ultimos_emergenciais = []

        start_handler = CommandHandler(command="start", callback=self.start)
        application.add_handler(start_handler)

        comandos_handler = CommandHandler(command="comandos", callback=self.comandos)
        application.add_handler(comandos_handler)

        application.add_handler(CommandHandler(command="monitorar_cada", callback=self.set_timer))

        application.add_handler(CommandHandler(command="monitorar", callback=self.set_timer))

        application.add_handler(CommandHandler(command="parar_monitoramento", callback=self.unset))

        application.add_handler(CommandHandler(command="parar", callback=self.unset))

        application.add_handler(CommandHandler(command="conferir_tela_21", callback=self.tela21))

        application.add_handler(CommandHandler(command="21", callback=self.tela21))

        application.add_handler(CommandHandler(
            command="21_programar", callback=self.tela21_programar
        ))

        application.add_handler(CommandHandler(
            command="21_", callback=self.tela21_programar
        ))

        application.add_handler(CommandHandler(command="enviar_para", callback=self.programar_p_equipe))

        application.add_handler(CommandHandler(command="32", callback=self.programar_p_equipe))

        application.add_handler(CommandHandler(command="imprimir", callback=self.imprimir_ase))

        application.add_handler(CommandHandler(command="58", callback=self.imprimir_ase))

        application.add_handler(CommandHandler(command="58_", callback=self.imprimir_ase_lista))

        application.add_handler(MessageHandler(filters.COMMAND, self.unknown))

        application.add_handler(
            MessageHandler(filters.TEXT & ~filters.COMMAND, self.echo)
        )

        application.run_polling()

    # BASIC FUNCTIONS

    async def start(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        while True:
            try:
                await context.bot.send_message(
                    chat_id=update.effective_chat.id,
                    text="Bot desenvolvido por Marcelo Santos para programação de serviços Sanepar\n\
                Envie /comandos para ver a lista de comandos disponiveis.",
                )
                break
            except:
                continue

    async def comandos(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        while True:
            try:
                await context.bot.send_message(
                    chat_id=update.effective_chat.id,
                    text="Os comandos disponiveis até o momento são:\n\
        /conferir_tela_21 \
        - Este comando deve sempre ser o primeiro a ser executado, \
        irá retornar todos os protocolos disponiveis na tela 21 emergencial.\n\n\
        /enviar_para <equipe> \
        - Deve ser sempre enviado como resposta para o protocolo ao qual deseja enviar, \
        exemplo: \n/enviar_para 2001 irá programar e enviar o protocolo para a equipe 2001.\n\n\
        /monitorar_cada <minutos> \
        - se intervalo não for informado utilizará 10 minutos como padrão \
        - irá verificar automaticamente a tela 21 a cada tanto tempo e retornar os protocolos disponiveis.\n\n\
        /parar_monitoramento \
        - ira parar a verificação automatica de protocolos na tela 21.\n\n\
        /imprimir <impressora> \
        - se não for informado utilizará <PRINTER> como padrão \
        - Deve ser sempre enviado como resposta para o protocolo ao qual deseja imprimir, \
        exemplo: \n/imprimir <PRINTER> irá enviar o Ase para a impressora <PRINTER>.",
                )
                break
            except:
                continue

    async def echo(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        if update.message.text == "tset":
            while True:
                try:
                    await context.bot.send_message(
                        chat_id=update.effective_chat.id, text="Test"
                    )
                    break
                except:
                    continue

    # FUNCTIONS SGC

    async def tela21(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        while True:
            try:
                await context.bot.send_message(
                    chat_id=update.effective_chat.id, text="Verificando tela 21, aguarde."
                )
                break
            except:
                continue
        main_sgc.sgc(chave=self.chave, senha=self.senha, impressora=self.impressora)
        self.dados, self.index = main_sgc.verificar_tela_21_emergencial()

        if len(self.dados) == 0:
            while True:
                try:
                    await context.bot.send_message(
                        chat_id=update.effective_chat.id, text="Sem protocolos pendentes"
                    )
                    return
                except:
                    continue
        else:
            while True:
                try:
                    await context.bot.send_message(
                        chat_id=update.effective_chat.id,
                        text="Protocolos atuais na tela 21 emergencial",
                    )
                    break
                except:
                    continue
            for i in range(0, len(self.dados), 6):
                corpo = f"\
    COD: {self.dados[i]}\n\
    Protocolo: {self.dados[i + 1][:8]}.{self.dados[i + 1][8:12]}.{self.dados[i + 1][12:]}\n\
    Equipe: {self.dados[i + 2]}\n\
    Local: {self.dados[i + 3]}\n\
    Endereço: {self.dados[i + 4]}\n\
    Motivo: {self.dados[i + 5]}"
            while True:
                try:
                    await context.bot.send_message(chat_id=update.effective_chat.id, text=corpo)
                    break
                except:
                    continue

    async def tela21_programar(
            self, update: Update, context: ContextTypes.DEFAULT_TYPE
    ):
        try:
            base = str(context.args[0])
        except:
            await update.effective_message.reply_text(
                'Comando invalido, use "/21_programar uva" para programar serviços da base de união da vitória'
            )
            return
        eqpes = []
        if base == "uva":
            eqpes = ['1999', "2999", "2999", "2999", "2999", "2999", "2999", "2999", "2999", "2999"]
            tipo = ['Equipe provisória: ',
                    'Equipe para HDs: ',
                    'Equipe caixa padrão subterrânea: ',
                    'Equipe caixa padrão de muro: ',
                    'Equipe para ligações: ',
                    'Equipe para manutenções: ',
                    'Equipe para calçadas:',
                    'Equipe para asfaltos:',
                    'Equipe para esgotos: ',
                    'Equipe para reaterros: '
                    ]
            for i in range(len(tipo)):
                print(tipo[i] + eqpes[i])
            resposta = input('confirma as equipes?\n')
            if resposta[:1] == 'n':
                eqpes = []
                for i in range(len(tipo)):
                    eqpes.append(input(f'{tipo[i]}:\n'))
        elif base == "bituruna":
            eqpes = ["1999", "2999", "2999", "2999", "2999", "2999", "2999", "2999", "2999", "2999"]
            tipo = ['Equipe provisória: ',
                    'Equipe para HDs: ',
                    'Equipe caixa padrão subterrânea: ',
                    'Equipe caixa padrão de muro: ',
                    'Equipe para ligações: ',
                    'Equipe para manutenções: ',
                    'Equipe para calçadas:',
                    'Equipe para asfaltos:',
                    'Equipe para esgotos: ',
                    'Equipe para reaterros: '
                    ]
            for i in range(len(tipo)):
                print(tipo[i] + eqpes[i])
            resposta = input('confirma as equipes?\n')
            if resposta[:1] == 'n':
                eqpes = []
                for i in range(len(tipo)):
                    eqpes.append(input(f'{tipo[i]}:\n'))

        await update.message.reply_text("Programando serviços, aguarde.")

        main_sgc.sgc(chave=self.chave, senha=self.senha, impressora=self.impressora)

        print(f"Equipes em main.py: {eqpes}")
        dados, excecoes = main_sgc.programar_servicos(equipes=eqpes, base=base)
        print("       SERV          PROTOCOLO                PREVIS            EQPE     LOC")
        linha = 0
        cod = []
        protocolos = []
        for i in range(0, len(dados), 5):
            linha += 1
            serv = dados[i]
            protocolo = f"{dados[i + 1][:8]}.{dados[i + 1][8:12]}.{dados[i + 1][12:]}"
            cod.append(serv)
            protocolos.append(dados[i + 1])
            data_vencimento = datetime.datetime.strptime(dados[i + 2], "%d%m%Y")
            equipe = dados[i + 3]
            local = dados[i + 4]
            print(
                f"{str(linha).zfill(3)}..: {serv}     {protocolo}     {data_vencimento}     {equipe}     {local}"
            )
        for i in range(0, len(excecoes), 2):
            print(f"Exceção.: {excecoes[i]}")
            print(f"linha.: {excecoes[i + 1]}")
        while True:
            try:
                await context.bot.send_message(
                    chat_id=update.effective_chat.id,
                    text="Programação completa.",
                )
                break
            except:
                continue
        resposta = input("Imprimir?\n")
        if resposta[:1].lower() == "s":
            for i in range(len(protocolos)):
                main_sgc.imprimir_servico(protocolo=protocolos[i], impressora=self.impressora)

    async def tela21_monitorar(self):
        main_sgc.sgc(chave=self.chave, senha=self.senha, impressora=self.impressora)
        self.dados, self.index = main_sgc.verificar_tela_21_emergencial()

    async def imprimir_ase(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        try:
            impressora = context.args[0]
            mensagem_pai = update.effective_message.reply_to_message.text
            protocolo = mensagem_pai[21:29] + mensagem_pai[30:34] + mensagem_pai[35:41]
            main_sgc.imprimir_servico(protocolo=protocolo, impressora=impressora)
            while True:
                try:
                    await context.bot.send_message(
                        chat_id=update.effective_chat.id,
                        text=f"Ase enviado para a impressora {self.impressora}",
                    )
                    break
                except:
                    continue
        except (IndexError, ValueError):
            impressora = "PDPGO125"
            mensagem_pai = update.effective_message.reply_to_message.text
            protocolo = mensagem_pai[21:29] + mensagem_pai[30:34] + mensagem_pai[35:41]
            main_sgc.imprimir_servico(protocolo=protocolo, impressora=impressora)
            while True:
                try:
                    await context.bot.send_message(
                        chat_id=update.effective_chat.id,
                        text=f"Ase enviado para a impressora {self.impressora}",
                    )
                    break
                except:
                    continue

    async def imprimir_ase_lista(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        protocolos = []
        impressora = "pdpgo125"
        while True:
            protocolo = input(f"Protocolo: ")
            if protocolo == "":
                break
            else:
                if len(protocolo) == 19:
                    protocolo = protocolo[:8] + protocolo[9:13] + protocolo[14:]
                    protocolos.append(protocolo)
                else:
                    protocolos.append(protocolo)
        for i in range(len(protocolos)):
            print(f"{i}: {protocolos[i]}")
            main_sgc.imprimir_servico(protocolo=protocolos[i], impressora=impressora)

    async def programar_p_equipe(
            self, update: Update, context: ContextTypes.DEFAULT_TYPE
    ):
        equipe = context.args[0]
        mensagem_pai = update.effective_message.reply_to_message.text
        protocolo = mensagem_pai[25:33] + mensagem_pai[34:38] + mensagem_pai[39:44]
        serv_cod = mensagem_pai[5:9]
        resposta = main_sgc.programar_servico_emergencial(
            codigo_serv=serv_cod, protocolo=protocolo, equipe=equipe)
        if not resposta:
            while True:
                try:
                    await context.bot.send_message(
                        chat_id=update.effective_chat.id,
                        text=f"Protocolo não encontrado"
                    )
                    return
                except:
                    continue
        resposta = main_sgc.enviar_serv_32(
            codigo_serv=serv_cod, protocolo=protocolo, equipe=equipe)
        if resposta:
            while True:
                try:
                    await context.bot.send_message(
                        chat_id=update.effective_chat.id,
                        text=(f"{mensagem_pai[25:33] + mensagem_pai[33:38] + mensagem_pai[38:44]} "
                            f"programado e enviado para a equipe {equipe}"),
                    )
                    break
                except:
                    continue
        if not resposta:
            while True:
                try:
                    await context.bot.send_message(
                        chat_id=update.effective_chat.id,
                        text=f"Protocolo não encontrado"
                    )
                    break
                except:
                    continue

    #  HELPERS

    async def set_timer(
            self, update: Update, context: ContextTypes.DEFAULT_TYPE
    ) -> None:
        """Add a job to the queue."""

        chat_id = update.effective_message.chat_id

        try:
            # args[0] should contain the time for the timer in minutes

            due = float(context.args[0]) * 60
            self.intervalo = float(context.args[0]) * 60
            self.monitorar = True

            if due < 300:
                await update.effective_message.reply_text(
                    "O intervalo não pode ser meno que 5 minutos"
                )
                return

            job_removed = self.remove_job_if_exists(str(chat_id), context)

            context.job_queue.run_repeating(
                callback=self.alarm,
                interval=due,
                first=1,
                chat_id=chat_id,
                name=str(chat_id),
                data=due,
            )

            text = f"Monitorando serviços emergenciais a cada {due / 60} minutos"

            if job_removed:
                text += "\nTempo de monitoramento alterado"

            await update.effective_message.reply_text(text)

        except (IndexError, ValueError):
            # await update.effective_message.reply_text("Uso: /monitorar_cada <minutos>")

            due = 10 * 60
            self.intervalo = 10 * 60
            self.monitorar = True

            job_removed = self.remove_job_if_exists(str(chat_id), context)

            context.job_queue.run_repeating(
                callback=self.alarm,
                interval=due,
                first=1,
                chat_id=chat_id,
                name=str(chat_id),
                data=due,
            )

            text = f"Monitorando serviços emergenciais a cada {due / 60} minutos"

            if job_removed:
                text += "Tempo de monitoramento alterado"

            await update.effective_message.reply_text(text)

    async def alarm(self, context: ContextTypes.DEFAULT_TYPE) -> None:
        """Send the alarm message."""

        job = context.job

        main_sgc.sgc(chave=self.chave, senha=self.senha, impressora=self.impressora)
        self.dados, self.index = main_sgc.verificar_tela_21_emergencial()

        if len(self.dados) == 0:
            while True:
                try:
                    await context.bot.send_message(
                        chat_id=job.chat_id, text="Sem protocolos pendentes"
                    )
                    return
                except:
                    continue
        elif self.dados == self.ultimos_emergenciais:
            while True:
                try:
                    await context.bot.send_message(
                        chat_id=job.chat_id, text="Sem novos protocolos"
                    )
                    return
                except:
                    continue
        else:
            self.ultimos_emergenciais = self.dados
            while True:
                try:
                    await context.bot.send_message(
                        chat_id=job.chat_id,
                        text="Protocolos atuais na tela 21 emergencial",
                    )
                    break
                except:
                    continue
            for i in range(0, len(self.dados), 6):
                corpo = f"\
    COD: {self.dados[i]}\n\
    Protocolo: {self.dados[i + 1][:8]}.{self.dados[i + 1][8:12]}.{self.dados[i + 1][12:]}\n\
    Equipe: {self.dados[i + 2]}\n\
    Local: {self.dados[i + 3]}\n\
    Endereço: {self.dados[i + 4]}\n\
    Motivo: {self.dados[i + 5]}"
            while True:
                try:
                    await context.bot.send_message(chat_id=job.chat_id, text=corpo)
                    break
                except:
                    continue

    def remove_job_if_exists(
            self, name: str, context: ContextTypes.DEFAULT_TYPE
    ) -> bool:
        """Remove job with given name. Returns whether job was removed."""

        current_jobs = context.job_queue.get_jobs_by_name(name)

        if not current_jobs:
            return False

        for job in current_jobs:
            job.schedule_removal()

        return True

    async def unset(self, update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
        """Remove the job if the user changed their mind."""

        chat_id = update.message.chat_id
        self.monitorar = False
        self.intervalo = 0

        job_removed = self.remove_job_if_exists(str(chat_id), context)

        text = (
            "Monitoramento desligado" if job_removed else "Monitoramento não esta ativo"
        )

        await update.message.reply_text(text)

    async def wait_for_response(
            self, context: CallbackContext, user_id: int
    ) -> Message:
        max_retries = 30
        retry_delay_seconds = 1

        for _ in range(max_retries):
            try:
                updates = await context.bot.get_updates(
                    allowed_updates=[Update.MESSAGE],
                )
                if updates and updates[-1].message.from_user.id == user_id:
                    return updates[-1].message
            except (TimeoutError, httpx.PoolTimeout):
                print("Request timed out. Retrying...")
                time.sleep(retry_delay_seconds)

    async def unknown(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        while True:
            try:
                await context.bot.send_message(
                    chat_id=update.effective_chat.id,
                    text="Comando não reconhecido.",
                )
                break
            except:
                continue


TOKEN = "<TOKEN>"

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s", level=logging.INFO
)

if __name__ == "__main__":
    app = app(token=TOKEN)
