import asyncio
import subprocess
import logging
import os
from pathlib import Path

logger = logging.getLogger(__name__)

class PdfConverterService:
    @staticmethod
    async def convert_to_pdf(pptx_path: str) -> str | None:
        """
        Converts a PPTX file to PDF using LibreOffice headless mode.
        Returns the path to the generated PDF file or None if failed.
        """
        path = Path(pptx_path)
        if not path.exists():
            logger.error(f"PPTX file not found for conversion: {pptx_path}")
            return None

        output_dir = path.parent
        pdf_path = path.with_suffix('.pdf')

        # Try common libreoffice binary names
        binaries = ['libreoffice', 'soffice']
        command_base = None
        
        for bin_name in binaries:
            try:
                # Check if binary exists in PATH
                check_process = await asyncio.create_subprocess_exec(
                    bin_name, '--version',
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE
                )
                await check_process.communicate()
                if check_process.returncode == 0:
                    command_base = bin_name
                    break
            except FileNotFoundError:
                continue

        if not command_base:
            logger.error("LibreOffice binary ('libreoffice' or 'soffice') not found in PATH.")
            logger.error(f"Current PATH: {os.environ.get('PATH')}")
            return None

        try:
            command = [
                command_base,
                '--headless',
                '--convert-to', 'pdf',
                '--outdir', str(output_dir),
                str(path)
            ]
            
            logger.info(f"Starting PDF conversion using {command_base}: {' '.join(command)}")
            
            # Run the command
            process = await asyncio.create_subprocess_exec(
                *command,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE
            )
            stdout, stderr = await process.communicate()
            
            if process.returncode == 0:
                if pdf_path.exists():
                    logger.info(f"PDF conversion successful: {pdf_path}")
                    return str(pdf_path)
                else:
                    logger.error("LibreOffice returned 0 but PDF file was not created.")
            else:
                logger.error(f"LibreOffice conversion failed with return code {process.returncode}")
                logger.error(f"Stdout: {stdout.decode()}")
                logger.error(f"Stderr: {stderr.decode()}")
                
        except Exception as e:
            logger.exception(f"Exception during PDF conversion: {e}")
            
        return None
