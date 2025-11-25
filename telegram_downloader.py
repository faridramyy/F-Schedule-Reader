import asyncio
import os
from telegram import Update
from telegram.ext import Application, MessageHandler, filters, ContextTypes

DOWNLOAD_DIR = "./downloads"
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

class XLSDownloader:
    def __init__(self, token: str):
        self.token = token
        # Future to store the result (path) once the file is downloaded
        self.file_future = None

    async def handle_document(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        """
        Handles the document upload.
        Only saves the file and updates the future. 
        Does NOT stop the application (we leave that to the main loop).
        """
        doc = update.message.document

        # Filter for XLS/XLSX
        if not doc.file_name.lower().endswith((".xls", ".xlsx")):
            await update.message.reply_text("❌ Please send an XLS/XLSX file.")
            return

        # Download file
        tg_file = await doc.get_file()
        save_path = f"{DOWNLOAD_DIR}/{doc.file_unique_id}_{doc.file_name}"
        await tg_file.download_to_drive(save_path)

        await update.message.reply_text("✅ File received! Closing connection...")

        # Signal that we are done by setting the result on the future
        if self.file_future and not self.file_future.done():
            self.file_future.set_result(save_path)

    async def start_and_wait(self):
        """
        Manually manages the bot lifecycle to allow for a clean exit
        after a specific condition (file download) is met.
        """
        # Get the current running loop
        loop = asyncio.get_running_loop()
        self.file_future = loop.create_future()

        # Build Application
        app = Application.builder().token(self.token).build()
        app.add_handler(MessageHandler(filters.Document.ALL, self.handle_document))

        print("🤖 Bot is waiting for your XLS file... Send it now!")

        # 1. Initialize and Start App
        await app.initialize()
        await app.start()

        # 2. Start Polling (Non-blocking)
        # allowed_updates ignores chat_member updates etc, saving bandwidth
        await app.updater.start_polling(allowed_updates=Update.ALL_TYPES)

        try:
            # 3. Block here until the file is received (Future resolves)
            file_path = await self.file_future
            print(f"📂 File captured: {file_path}")
            return file_path

        except Exception as e:
            print(f"⚠️ Error during waiting: {e}")
            raise e

        finally:
            # 4. Clean Shutdown Sequence (CRITICAL for avoiding RuntimeErrors)
            print("🛑 Shutting down bot...")
            
            # Stop the updater first (stops fetching new updates)
            if app.updater.running:
                await app.updater.stop()
            
            # Stop the application (stops processing current updates)
            if app.running:
                await app.stop()
            
            # Shutdown (closes http sessions and releases resources)
            await app.shutdown()
            print("👋 Bot shutdown complete.")

async def _run_downloader(token: str):
    downloader = XLSDownloader(token)
    return await downloader.start_and_wait()

def wait_for_xls(token: str):
    """
    Entry point.
    """
    try:
        # Check if there is already a running loop (e.g., inside Jupyter or existing asyncio app)
        loop = asyncio.get_running_loop()
    except RuntimeError:
        loop = None

    if loop and loop.is_running():
        # If we are already in an async environment, return the coroutine
        # Note: This requires the caller to await this function, 
        # but typical synchronous usage (like main.py) usually uses asyncio.run
        print("⚠️ Warning: An event loop is already running. Returning coroutine.")
        return _run_downloader(token)
    else:
        # Standard usage: Create a new loop
        return asyncio.run(_run_downloader(token))