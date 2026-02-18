"""Tracker module for processing CSV files to Excel.

Commands:
    !tracker upload           - Interactive upload wizard
    !tracker upload master    - Upload master roster CSV
    !tracker upload typeform  - Upload typeform responses CSV
    !tracker upload zoom      - Upload zoom attendance CSV
    !tracker download         - Generate Excel report from uploaded CSVs
    !tracker files            - Show status of uploaded CSV files
    !tracker clear <type>     - Clear specific CSV file
    !tracker clearall         - Clear all uploaded CSV files
    !tracker help             - Show help (handled by bot/events.py)
"""

import asyncio
import io
from typing import Optional

import discord
from discord.ext import commands

from services.file_processor import FileStorageService, VALID_FILE_CATEGORIES
from services.tracker_processor import TrackerDataProcessor


# File category descriptions
FILE_DESCRIPTIONS = {
    "master": "Master Roster (student list with enrollment data)",
    "typeform": "Typeform Responses (weekly progress submissions)",
    "zoom": "Zoom Attendance (lecture/office hours attendance)"
}


class TrackerCog(commands.Cog, name="Tracker"):
    """Cog for processing tracker CSV files.
    
    Supports uploading 3 separate CSV files (master, typeform, zoom)
    and generating comprehensive Excel reports.
    
    Note: Bot uses '!tracker ' as prefix, so commands are direct (not subcommands).
    """
    
    def __init__(self, bot: commands.Bot):
        self.bot = bot
        self.storage = FileStorageService()
        self.processor = TrackerDataProcessor()
        # Track users in upload wizard to prevent conflicts
        self._upload_sessions: dict[int, str] = {}
    
    async def _wait_for_csv(self, ctx: commands.Context, 
                           category: str, timeout: float = 120.0) -> Optional[bytes]:
        """Wait for a CSV file upload from the user.
        
        Returns the file bytes if successful, None if cancelled or timed out.
        """
        def check(message: discord.Message) -> bool:
            # Same user, same channel
            if message.author.id != ctx.author.id or message.channel.id != ctx.channel.id:
                return False
            
            # Check for cancel command
            if message.content.lower() in ['cancel', '!cancel']:
                return True
            
            # Check for CSV attachment
            for attachment in message.attachments:
                if attachment.filename.lower().endswith('.csv'):
                    return True
            
            return False
        
        try:
            message = await self.bot.wait_for('message', check=check, timeout=timeout)
            
            # Check if cancelled
            if message.content.lower() in ['cancel', '!cancel']:
                return None
            
            # Get CSV attachment
            for attachment in message.attachments:
                if attachment.filename.lower().endswith('.csv'):
                    file_data = await attachment.read()
                    
                    # Store the file
                    stored_file = self.storage.store_file(
                        filename=attachment.filename,
                        data=file_data,
                        user_id=ctx.author.id,
                        category=category
                    )
                    
                    # Format file size
                    size_kb = len(file_data) / 1024
                    size_str = f"{size_kb:.1f} KB" if size_kb < 1024 else f"{size_kb/1024:.1f} MB"
                    
                    await ctx.send(
                        f"✅ **{category.title()} CSV Stored!**\n"
                        f"• File: `{attachment.filename}`\n"
                        f"• Size: {size_str}"
                    )
                    
                    return file_data
            
            return None
            
        except asyncio.TimeoutError:
            await ctx.send(f"⏱️ Upload timed out for {category} CSV.")
            return None
    
    @commands.command(name='files')
    async def files(self, ctx: commands.Context):
        """Show status of all uploaded CSV files."""
        files = self.storage.get_all_files()
        
        status_lines = ["**📁 Tracker CSV Status**\n"]
        
        for category in VALID_FILE_CATEGORIES:
            stored = files.get(category)
            desc = FILE_DESCRIPTIONS.get(category, category)
            
            if stored:
                # Format upload time
                upload_time = stored.uploaded_at.strftime("%Y-%m-%d %H:%M")
                status_lines.append(
                    f"✅ **{category.title()}** ({desc})\n"
                    f"   └─ `{stored.filename}` (uploaded {upload_time})"
                )
            else:
                status_lines.append(
                    f"❌ **{category.title()}** ({desc})\n"
                    f"   └─ Not uploaded"
                )
        
        await ctx.send("\n".join(status_lines))
    
    @commands.group(name='upload', invoke_without_command=True)
    async def upload(self, ctx: commands.Context):
        """Interactive upload wizard - prompts for each CSV file."""
        # Check if user already in upload session
        if ctx.author.id in self._upload_sessions:
            await ctx.send("⚠️ You already have an upload session in progress.")
            return
        
        self._upload_sessions[ctx.author.id] = "wizard"
        
        try:
            await ctx.send(
                "**📤 Tracker Upload Wizard**\n\n"
                "I'll guide you through uploading each CSV file.\n"
                "For each file, you can:\n"
                "• Upload a CSV file\n"
                "• Type `skip` to skip that file\n"
                "• Type `cancel` to abort the wizard\n"
                "─────────────────────────────"
            )
            
            for category in ["master", "typeform", "zoom"]:
                desc = FILE_DESCRIPTIONS.get(category, category)
                existing = self.storage.get_file(category)
                
                existing_info = ""
                if existing:
                    existing_info = f"\n   └─ Current: `{existing.filename}`"
                
                await ctx.send(
                    f"\n**{category.upper()}** - {desc}{existing_info}\n"
                    f"Upload the {category} CSV file, type `skip`, or type `cancel`:"
                )
                
                # Wait for response
                def check(message: discord.Message) -> bool:
                    if message.author.id != ctx.author.id or message.channel.id != ctx.channel.id:
                        return False
                    
                    content = message.content.lower().strip()
                    if content in ['skip', 'cancel', '!cancel']:
                        return True
                    
                    for attachment in message.attachments:
                        if attachment.filename.lower().endswith('.csv'):
                            return True
                    
                    return False
                
                try:
                    message = await self.bot.wait_for('message', check=check, timeout=120.0)
                    content = message.content.lower().strip()
                    
                    if content in ['cancel', '!cancel']:
                        await ctx.send("❌ Upload wizard cancelled.")
                        return
                    
                    if content == 'skip':
                        await ctx.send(f"⏭️ Skipped {category} CSV.")
                        continue
                    
                    # Process CSV upload
                    for attachment in message.attachments:
                        if attachment.filename.lower().endswith('.csv'):
                            file_data = await attachment.read()
                            
                            self.storage.store_file(
                                filename=attachment.filename,
                                data=file_data,
                                user_id=ctx.author.id,
                                category=category
                            )
                            
                            size_kb = len(file_data) / 1024
                            size_str = f"{size_kb:.1f} KB" if size_kb < 1024 else f"{size_kb/1024:.1f} MB"
                            
                            await ctx.send(
                                f"✅ **{category.title()} CSV Stored!**\n"
                                f"   • File: `{attachment.filename}`\n"
                                f"   • Size: {size_str}"
                            )
                            break
                    
                except asyncio.TimeoutError:
                    await ctx.send(f"⏱️ Timed out waiting for {category} CSV. Wizard ended.")
                    return
            
            # Wizard complete
            await ctx.send(
                "─────────────────────────────\n"
                "**✅ Upload Wizard Complete!**\n\n"
                "Run `!tracker files` to see all uploaded files.\n"
                "Run `!tracker download` to generate the report."
            )
            
        finally:
            # Clean up session
            self._upload_sessions.pop(ctx.author.id, None)
    
    @upload.command(name='master')
    async def upload_master(self, ctx: commands.Context):
        """Upload master roster CSV file."""
        existing = self.storage.get_file("master")
        existing_info = f"\n   └─ Current: `{existing.filename}`" if existing else ""
        
        await ctx.send(
            f"**📤 Upload Master Roster CSV**{existing_info}\n\n"
            f"Please upload the master roster CSV file, or type `cancel` to abort:"
        )
        
        await self._wait_for_csv(ctx, "master")
    
    @upload.command(name='typeform')
    async def upload_typeform(self, ctx: commands.Context):
        """Upload typeform responses CSV file."""
        existing = self.storage.get_file("typeform")
        existing_info = f"\n   └─ Current: `{existing.filename}`" if existing else ""
        
        await ctx.send(
            f"**📤 Upload Typeform Responses CSV**{existing_info}\n\n"
            f"Please upload the typeform responses CSV file, or type `cancel` to abort:"
        )
        
        await self._wait_for_csv(ctx, "typeform")
    
    @upload.command(name='zoom')
    async def upload_zoom(self, ctx: commands.Context):
        """Upload zoom attendance CSV file."""
        existing = self.storage.get_file("zoom")
        existing_info = f"\n   └─ Current: `{existing.filename}`" if existing else ""
        
        await ctx.send(
            f"**📤 Upload Zoom Attendance CSV**{existing_info}\n\n"
            f"Please upload the zoom attendance CSV file, or type `cancel` to abort:"
        )
        
        await self._wait_for_csv(ctx, "zoom")
    
    @commands.group(name='clear', invoke_without_command=True)
    async def clear(self, ctx: commands.Context):
        """Clear uploaded CSV files. Use subcommands to specify which file."""
        await ctx.send(
            "**🗑️ Clear CSV Files**\n\n"
            "Use one of the following commands:\n"
            "• `!tracker clear master` - Remove master roster CSV\n"
            "• `!tracker clear typeform` - Remove typeform responses CSV\n"
            "• `!tracker clear zoom` - Remove zoom attendance CSV\n"
            "• `!tracker clearall` - Remove all CSV files"
        )
    
    @clear.command(name='master')
    async def clear_master(self, ctx: commands.Context):
        """Clear the master roster CSV file."""
        if self.storage.delete_file("master"):
            await ctx.send("✅ **Master CSV cleared!**")
        else:
            await ctx.send("ℹ️ No master CSV file to clear.")
    
    @clear.command(name='typeform')
    async def clear_typeform(self, ctx: commands.Context):
        """Clear the typeform responses CSV file."""
        if self.storage.delete_file("typeform"):
            await ctx.send("✅ **Typeform CSV cleared!**")
        else:
            await ctx.send("ℹ️ No typeform CSV file to clear.")
    
    @clear.command(name='zoom')
    async def clear_zoom(self, ctx: commands.Context):
        """Clear the zoom attendance CSV file."""
        if self.storage.delete_file("zoom"):
            await ctx.send("✅ **Zoom CSV cleared!**")
        else:
            await ctx.send("ℹ️ No zoom CSV file to clear.")
    
    @commands.command(name='clearall')
    async def clearall(self, ctx: commands.Context):
        """Clear all uploaded CSV files."""
        deleted = self.storage.delete_all_files()
        if deleted > 0:
            await ctx.send(f"✅ **All CSV files cleared!** ({deleted} file(s) removed)")
        else:
            await ctx.send("ℹ️ No CSV files to clear.")
    
    @commands.command(name='download')
    async def download(self, ctx: commands.Context):
        """Process uploaded CSV files and return a styled Excel file.
        
        Usage:
            1. Upload CSV files using !tracker upload commands
            2. Run !tracker download to generate the report
        """
        # Get the typeform file (primary data source)
        typeform_file = self.storage.get_file("typeform")
        
        if typeform_file is None:
            await ctx.send(
                "❌ **No typeform CSV uploaded.**\n\n"
                "The typeform CSV is required for generating reports.\n"
                "Upload it using `!tracker upload typeform`."
            )
            return
        
        # Check for optional files
        master_file = self.storage.get_file("master")
        zoom_file = self.storage.get_file("zoom")
        
        files_info = [f"• Typeform: `{typeform_file.filename}`"]
        if master_file:
            files_info.append(f"• Master: `{master_file.filename}`")
        if zoom_file:
            files_info.append(f"• Zoom: `{zoom_file.filename}`")
        
        await ctx.send(
            f"📂 **Processing Files:**\n" + "\n".join(files_info) + 
            "\n\n⏳ Creating multi-tab report..."
        )
        
        # Process the files
        try:
            # Read all available files
            typeform_data = self.storage.read_file(typeform_file)
            master_data = self.storage.read_file(master_file) if master_file else None
            zoom_data = self.storage.read_file(zoom_file) if zoom_file else None
            
            # Process with tracker processor (pass all data sources)
            result = self.processor.process(
                typeform_data,
                options={
                    'master_data': master_data,
                    'zoom_data': zoom_data
                }
            )
            
            if not result.success:
                await ctx.send(f"❌ Processing failed: {result.error_message}")
                return
            
            # Generate output filename
            base_name = typeform_file.filename.rsplit('.', 1)[0]
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
