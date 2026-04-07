param(
    [double]$Scale = 1.0,
    [int]$YOffset = 0,
    [string]$PngPath = (Join-Path $PSScriptRoot 'AppIcon.png'),
    [string]$IcoPath = (Join-Path $PSScriptRoot 'AppIcon.ico')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Drawing

$code = @"
using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;

public static class PassageIconTool
{
    public static Rectangle GetOpaqueBounds(Bitmap bitmap)
    {
        int minX = bitmap.Width;
        int minY = bitmap.Height;
        int maxX = -1;
        int maxY = -1;

        for (int y = 0; y < bitmap.Height; y++)
        {
            for (int x = 0; x < bitmap.Width; x++)
            {
                if (bitmap.GetPixel(x, y).A <= 8)
                {
                    continue;
                }

                if (x < minX) minX = x;
                if (y < minY) minY = y;
                if (x > maxX) maxX = x;
                if (y > maxY) maxY = y;
            }
        }

        if (maxX < minX || maxY < minY)
        {
            return Rectangle.Empty;
        }

        return Rectangle.FromLTRB(minX, minY, maxX + 1, maxY + 1);
    }

    public static void TransformPng(string inputPath, string outputPath, double scale, int yOffset)
    {
        using (var source = new Bitmap(inputPath))
        {
            Rectangle bounds = GetOpaqueBounds(source);
            if (bounds.Width <= 0 || bounds.Height <= 0)
            {
                bounds = new Rectangle(0, 0, source.Width, source.Height);
            }

            double centerX = bounds.Left + (bounds.Width / 2.0);
            double centerY = bounds.Top + (bounds.Height / 2.0);
            int scaledWidth = Math.Max(1, (int)Math.Round(bounds.Width * scale));
            int scaledHeight = Math.Max(1, (int)Math.Round(bounds.Height * scale));
            int targetX = (int)Math.Round(centerX - (scaledWidth / 2.0));
            int targetY = (int)Math.Round(centerY - (scaledHeight / 2.0) + yOffset);

            int maxX = Math.Max(0, source.Width - scaledWidth);
            int maxY = Math.Max(0, source.Height - scaledHeight);
            targetX = Math.Max(0, Math.Min(maxX, targetX));
            targetY = Math.Max(0, Math.Min(maxY, targetY));

            using (var output = new Bitmap(source.Width, source.Height, PixelFormat.Format32bppArgb))
            using (var graphics = Graphics.FromImage(output))
            {
                graphics.CompositingMode = CompositingMode.SourceOver;
                graphics.CompositingQuality = CompositingQuality.HighQuality;
                graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
                graphics.SmoothingMode = SmoothingMode.HighQuality;
                graphics.Clear(Color.Transparent);
                graphics.DrawImage(
                    source,
                    new Rectangle(targetX, targetY, scaledWidth, scaledHeight),
                    bounds,
                    GraphicsUnit.Pixel);
                output.Save(outputPath, ImageFormat.Png);
            }
        }
    }

    public static void BuildIco(string pngPath, string icoPath, int[] sizes)
    {
        using (var source = new Bitmap(pngPath))
        using (var prepared = CreateSquareCanvas(source))
        using (var output = new FileStream(icoPath, FileMode.Create, FileAccess.Write))
        using (var writer = new BinaryWriter(output))
        {
            var imageData = new byte[sizes.Length][];
            for (int i = 0; i < sizes.Length; i++)
            {
                int size = sizes[i];
                using (var bitmap = new Bitmap(size, size, PixelFormat.Format32bppArgb))
                using (var graphics = Graphics.FromImage(bitmap))
                using (var stream = new MemoryStream())
                {
                    graphics.CompositingMode = CompositingMode.SourceOver;
                    graphics.CompositingQuality = CompositingQuality.HighQuality;
                    graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                    graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
                    graphics.SmoothingMode = SmoothingMode.HighQuality;
                    graphics.Clear(Color.Transparent);
                    graphics.DrawImage(prepared, new Rectangle(0, 0, size, size), new Rectangle(0, 0, prepared.Width, prepared.Height), GraphicsUnit.Pixel);
                    bitmap.Save(stream, ImageFormat.Png);
                    imageData[i] = stream.ToArray();
                }
            }

            writer.Write((short)0);
            writer.Write((short)1);
            writer.Write((short)imageData.Length);

            int offset = 6 + (16 * imageData.Length);
            for (int i = 0; i < sizes.Length; i++)
            {
                int size = sizes[i];
                byte dim = size >= 256 ? (byte)0 : (byte)size;
                writer.Write(dim);
                writer.Write(dim);
                writer.Write((byte)0);
                writer.Write((byte)0);
                writer.Write((short)1);
                writer.Write((short)32);
                writer.Write(imageData[i].Length);
                writer.Write(offset);
                offset += imageData[i].Length;
            }

            for (int i = 0; i < imageData.Length; i++)
            {
                writer.Write(imageData[i]);
            }
        }
    }

    private static Bitmap CreateSquareCanvas(Bitmap source)
    {
        int side = Math.Max(source.Width, source.Height);
        var canvas = new Bitmap(side, side, PixelFormat.Format32bppArgb);
        using (var graphics = Graphics.FromImage(canvas))
        {
            graphics.CompositingMode = CompositingMode.SourceOver;
            graphics.CompositingQuality = CompositingQuality.HighQuality;
            graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
            graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
            graphics.SmoothingMode = SmoothingMode.HighQuality;
            graphics.Clear(Color.Transparent);

            int x = (side - source.Width) / 2;
            int y = (side - source.Height) / 2;
            graphics.DrawImage(source, new Rectangle(x, y, source.Width, source.Height), new Rectangle(0, 0, source.Width, source.Height), GraphicsUnit.Pixel);
        }

        return canvas;
    }
}
"@

Add-Type -TypeDefinition $code -ReferencedAssemblies System.Drawing

$tempPng = [System.IO.Path]::Combine([System.IO.Path]::GetDirectoryName($PngPath), ([System.IO.Path]::GetFileNameWithoutExtension($PngPath) + '.tmp.png'))
$tempIco = [System.IO.Path]::Combine([System.IO.Path]::GetDirectoryName($IcoPath), ([System.IO.Path]::GetFileNameWithoutExtension($IcoPath) + '.tmp.ico'))

if ($Scale -ne 1.0 -or $YOffset -ne 0) {
    [PassageIconTool]::TransformPng($PngPath, $tempPng, $Scale, $YOffset)
    Move-Item -LiteralPath $tempPng -Destination $PngPath -Force
}

[PassageIconTool]::BuildIco($PngPath, $tempIco, @(16, 20, 24, 32, 40, 48, 64, 128, 256))
Move-Item -LiteralPath $tempIco -Destination $IcoPath -Force

$img = [System.Drawing.Bitmap]::FromFile($PngPath)
try {
    $bounds = [PassageIconTool]::GetOpaqueBounds($img)
    Write-Output ("Updated {0}" -f $PngPath)
    Write-Output ("Scale={0} YOffset={1}" -f $Scale, $YOffset)
    Write-Output ("Canvas={0}x{1}" -f $img.Width, $img.Height)
    Write-Output ("OpaqueBounds={0},{1} to {2},{3}" -f $bounds.Left, $bounds.Top, ($bounds.Right - 1), ($bounds.Bottom - 1))
} finally {
    $img.Dispose()
}
