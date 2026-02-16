"""Tracker module for processing CSV files to Excel.

Commands:
    !tracker download - Process the last uploaded CSV file
    !tracker help     - Show help for tracker commands
    
Auto-detection:
    Upload a CSV file (no command needed) and it will be stored for processing.
"""

import io

import discord
from discord.ext import commands

from services.file_processor import FileStorageService
from services.tracker_processor import TrackerDataProcessor


class TrackerCog(commands.Cog, name="Tracker"):
    """Cog for processing tracker CSV files.
    
    Single Responsibility: Handles Discord command interface for file processing.
    Depends on abstractions (ProcessorRegistry, FileStorageService) not concretions.
    """
    
    def __init__(self, bot: commands.Bot):
        self.bot = bot
        self.storage = FileStorageService()
        self.processor = TrackerDataProcessor()
    
    @commands.Cog.listener()
    async def on_message(self, message: discord.Message):
        """Listen for CSV file uploads and store them automatically."""
        # Ignore bot messages
        if message.author.bot:
            return
        
        # Check for CSV attachments
        for attachment in message.attachments:
            if attachment.filename.lower().endswith('.csv'):
                try:
                    # Download and store the CSV
                    file_data = await attachment.read()
                    
                    stored_file = self.storage.store_file(
                        filename=attachment.filename,
                        data=file_data,
                        user_id=message.author.id
                    )
                    
                    # Format file size
                    size_kb = len(file_data) / 1024
                    size_str = f"{size_kb:.1f} KB" if size_kb < 1024 else f"{size_kb/1024:.1f} MB"
                    
                    await message.channel.send(
                        f"✅ **CSV Received!**\n"
                        f"• File: `{attachment.filename}`\n"
                        f"• Size: {size_str}\n"
                        f"• Ready for processing\n\n"
                        f"Run `!tracker download` to convert to Excel."
                    )
                    
                except Exception as e:
                    await message.channel.send(f"❌ Failed to store CSV: {e}")
                
                # Only process first CSV if multiple attached
                break
    
    @commands.command(name='download')
    async def download(self, ctx: commands.Context):
        """Process the last uploaded CSV file and return a styled Excel file.
        
        Usage:
            1. Upload a CSV file (no command needed)
            2. Run !tracker download to process it
        """
        # Get the last uploaded file
        stored_file = self.storage.get_last_file()
        
        if stored_file is None:
            await ctx.send("❌ No input CSVs provided.\n\nUpload a CSV file first, then run `!tracker download`.")
            return
        
        await ctx.send(f"📂 Processing: `{stored_file.filename}`...\n⏳ Creating multi-tab report...")
        
        # Process the file
        try:
            file_data = self.storage.read_file(stored_file)
            
            # Process with tracker processor
            result = self.processor.process(file_data)
            
            if not result.success:
                await ctx.send(f"❌ Processing failed: {result.error_message}")
                return
            
            # Generate output filename
            base_name = stored_file.filename.rsplit('.', 1)[0]
            output_filename = f"{base_name}_report.xlsx"
            
            # Create file from bytes and send
            file = discord.File(
                fp=io.BytesIO(result.output_data),
                filename=output_filename
            )
            
            await ctx.send(
                f"✅ **Tracker Report Generated!**\n"
                f"• Students processed: {result.rows_processed}\n"
                f"• Tabs created:\n"
                f"  └─ Master Tracker (all fields)\n"
                f"  └─ P1 - At Risk (red/orange/yellow coding)\n"
                f"  └─ P2 - Flagged (yellow coding)\n"
                f"  └─ P3 - On Track (green coding)\n"
                f"  └─ Weekly Summary (dashboard)",
                file=file
            )
            
        except Exception as e:
            await ctx.send(f"❌ Error processing file: {e}")


async def setup(bot: commands.Bot):
    """Setup function for loading the cog."""
    await bot.add_cog(TrackerCog(bot))

