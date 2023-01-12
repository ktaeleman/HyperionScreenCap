using HyperionScreenCap.Capture;
using Vortice;
using Vortice.Direct3D;
using Vortice.Direct3D11;
using Vortice.DXGI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using static Vortice.Direct3D11.D3D11;
using static Vortice.DXGI.DXGI;



namespace HyperionScreenCap
{

    static class MathExt
    {
        public static T Clamp<T>(this T val, T min, T max) where T : IComparable<T>
        {
            if (val.CompareTo(min) < 0) return min;
            else if (val.CompareTo(max) > 0) return max;
            else return val;
        }
    }

    class DX11ScreenCapture : IScreenCapture
    {
        private int _adapterIndex;
        private int _monitorIndex;
        private int _scalingFactor;
        private int _maxFps;
        private int _frameCaptureTimeout;

        private IDXGIFactory2 _factory;
        private IDXGIAdapter1 _adapter;
        private IDXGIOutput _output;
        private IDXGIOutput5 _output5;
        private ID3D11Device _device;
        private ID3D11Texture2D _stagingTexture;
        private ID3D11Texture2D _smallerTexture;
        private ID3D11ShaderResourceView _smallerTextureView;
        private IDXGIOutputDuplication _duplicatedOutput;
        private int _scalingFactorLog2;
        private int _width;
        private int _height;
        private byte[] _lastCapturedFrame;
        private int _minCaptureTime;
        private Stopwatch _captureTimer;
        private bool _desktopDuplicatorInvalid;
        private bool _disposed;
        private static readonly FeatureLevel[] s_featureLevels = new[]
        {
            FeatureLevel.Level_11_0
        };

        public int CaptureWidth { get; private set; }
        public int CaptureHeight { get; private set; }

        public static String GetAvailableMonitors()
        {
            StringBuilder response = new StringBuilder();
            IDXGIFactory2 tempFactory = null;
            DXGI.CreateDXGIFactory2(false,out tempFactory);
            //using ( IDXGIFactory2 factory = new IDXGIFactory2() )
            {
                int adapterIndex = 0;
                while(tempFactory.EnumAdapters(adapterIndex, out IDXGIAdapter adapter) == SharpGen.Runtime.Result.Ok)
                {
                    response.Append($"Adapter Index {adapterIndex++}: {adapter.Description.Description}\n");
                    int outputIndex = 0;
                    while(adapter.EnumOutputs(outputIndex, out IDXGIOutput output) == SharpGen.Runtime.Result.Ok)
                    {
                        response.Append($"\tMonitor Index {outputIndex++}: {output.Description.DeviceName}");
                        var desktopBounds = output.Description.DesktopCoordinates;
                        response.Append($" {desktopBounds.Right - desktopBounds.Left}×{desktopBounds.Bottom - desktopBounds.Top}\n");
                    }
                    response.Append("\n");
                }
            }
            tempFactory.Dispose();
            return response.ToString();
        }

        public DX11ScreenCapture(int adapterIndex, int monitorIndex, int scalingFactor, int maxFps, int frameCaptureTimeout)
        {
            _adapterIndex = adapterIndex;
            _monitorIndex = monitorIndex;
            _scalingFactor = scalingFactor;
            _maxFps = maxFps;
            _frameCaptureTimeout = frameCaptureTimeout;
            _disposed = true;
        }

        public void Initialize()
        {

            // Create DXGI IDXGIFactory2
            DXGI.CreateDXGIFactory2(false,out _factory);
            _factory.EnumAdapters1(_adapterIndex, out _adapter);

            // Create device from Adapter
            D3D11CreateDevice(_adapter, DriverType.Unknown, DeviceCreationFlags.BgraSupport, s_featureLevels, out _device);

            // Get DXGI.Output
            _adapter.EnumOutputs(_monitorIndex, out _output);
            _output5 = _output.QueryInterface<IDXGIOutput5>();

            // Width/Height of desktop to capture
            var desktopBounds = _output.Description.DesktopCoordinates;
            _width = desktopBounds.Right - desktopBounds.Left;
            _height = desktopBounds.Bottom - desktopBounds.Top;

            CaptureWidth = _width / _scalingFactor;
            CaptureHeight = _height / _scalingFactor;
            
            // Initialize duplicator so we can see what the output format is
            InitDesktopDuplicator();


            _minCaptureTime = 1000 / _maxFps;
            _captureTimer = new Stopwatch();
            _disposed = false;
        }

        private void InitDesktopDuplicator()
        {
            // We're potentially reinitializing the duplicator, which could change output format
            // So make sure we reinitialize our textures
            _stagingTexture?.Dispose();
            _smallerTexture?.Dispose();
            _smallerTextureView?.Dispose();

            // Duplicate the output
            Format[] DesktopFormats = {  Format.R16G16B16A16_Float, Format.B8G8R8A8_UNorm };
            _duplicatedOutput = _output5.DuplicateOutput1(_device, DesktopFormats);

            // Calculate miplevels
            int mipLevels;
            if ( _scalingFactor == 1 )
                mipLevels = 1;
            else if ( _scalingFactor > 0 && _scalingFactor % 2 == 0 )
            {
                /// Mip level for a scaling factor other than one is computed as follows:
                /// 2^n = 2 + n - 1 where LHS is the scaling factor and RHS is the MipLevels value.
                _scalingFactorLog2 = Convert.ToInt32(Math.Log(_scalingFactor, 2));
                mipLevels = 2 + _scalingFactorLog2 - 1;
            }
            else
                throw new Exception("Invalid scaling factor. Allowed valued are 1, 2, 4, etc.");

            // Create Staging texture CPU-accessible
            var stagingTextureDesc = new Texture2DDescription
            {
                CPUAccessFlags = CpuAccessFlags.Read,
                BindFlags = BindFlags.None,
                Format = _duplicatedOutput.Description.ModeDescription.Format,
                Width = CaptureWidth,
                Height = CaptureHeight,
                MiscFlags = ResourceOptionFlags.None,
                MipLevels = 1,
                ArraySize = 1,
                SampleDescription = { Count = 1, Quality = 0 },
                Usage = ResourceUsage.Staging
            };
            _stagingTexture = _device.CreateTexture2D(stagingTextureDesc);

            // Create smaller texture to downscale the captured image
            var smallerTextureDesc = new Texture2DDescription
            {
                CPUAccessFlags = CpuAccessFlags.None,
                BindFlags = BindFlags.RenderTarget | BindFlags.ShaderResource,
                Format = _duplicatedOutput.Description.ModeDescription.Format,
                Width = _width,
                Height = _height,
                MiscFlags = ResourceOptionFlags.GenerateMips,
                MipLevels = mipLevels,
                ArraySize = 1,
                SampleDescription = { Count = 1, Quality = 0 },
                Usage = ResourceUsage.Default
            };
            _smallerTexture = _device.CreateTexture2D(smallerTextureDesc);
            _smallerTextureView = _device.CreateShaderResourceView(_smallerTexture);

            _desktopDuplicatorInvalid = false;
        }

        public byte[] Capture()
        {
            if ( _desktopDuplicatorInvalid )
            {
                _duplicatedOutput?.Dispose();
                InitDesktopDuplicator();
            }

            _captureTimer.Restart();
            byte[] response = ManagedCapture();
            _captureTimer.Stop();

            return response;
        }

        private byte[] ManagedCapture()
        {
            IDXGIResource screenResource = null;
            OutduplFrameInfo duplicateFrameInformation;

            try
            {
                {
                    // Try to get duplicated frame within given time
                    SharpGen.Runtime.Result res = _duplicatedOutput.AcquireNextFrame(_frameCaptureTimeout, out duplicateFrameInformation, out screenResource);

                    if ( res == SharpGen.Runtime.Result.WaitTimeout.Code && _lastCapturedFrame != null )
                        return _lastCapturedFrame;

                    if (res == new SharpGen.Runtime.Result(0x887A0026)) // ACCESS_LOST
                    {
                        _desktopDuplicatorInvalid = true;
                        return null;
                    }
                    if ( duplicateFrameInformation.LastPresentTime == 0 && _lastCapturedFrame != null )
                        return _lastCapturedFrame;
                }

                // Check if scaling is used
                if ( CaptureWidth != _width )
                {
                    // Copy resource into memory that can be accessed by the CPU
                    using ( var screenTexture2D = screenResource.QueryInterface<ID3D11Texture2D>() )
                        _device.ImmediateContext.CopySubresourceRegion(_smallerTexture, 0, 0, 0, 0, screenTexture2D, 0);

                    // Generates the mipmap of the screen
                    _device.ImmediateContext.GenerateMips(_smallerTextureView);

                    // Copy the mipmap of smallerTexture (size/ scalingFactor) to the staging texture: 1 for /2, 2 for /4...etc
                    _device.ImmediateContext.CopySubresourceRegion(_stagingTexture, 0, 0, 0, 0, _smallerTexture, _scalingFactorLog2);
                }
                else
                {
                    // Copy resource into memory that can be accessed by the CPU
                    using ( var screenTexture2D = screenResource.QueryInterface<ID3D11Texture2D>() )
                        _device.ImmediateContext.CopyResource(screenTexture2D, _stagingTexture);
                }

                // Get the desktop capture texture
                MappedSubresource mapSource = _device.ImmediateContext.Map(_stagingTexture, 0, MapMode.Read);
                _lastCapturedFrame = ToRGBArray(mapSource, _stagingTexture.Description.Format);
                return _lastCapturedFrame;
            }
            finally
            {
                screenResource?.Dispose();
                // Fixed OUT_OF_MEMORY issue on AMD Radeon cards. Ignoring all exceptions during unmapping.
                try { _device.ImmediateContext.Unmap(_stagingTexture, 0); } catch { };
                // Ignore DXGI_ERROR_INVALID_CALL, DXGI_ERROR_ACCESS_LOST errors since capture is already complete
                try { _duplicatedOutput.ReleaseFrame(); } catch { }
            }
        }

        public static float Parse16BitFloat(byte Hi, byte Lo)
        {
            int fullFloat = ((Hi << 8) | Lo);
            int mant = fullFloat & 0x03ff;            // 10 bits mantissa
            int exp =  fullFloat & 0x7c00;            // 5 bits exponent
            if( exp == 0x7c00 )                   // NaN/Inf
                exp = 0x3fc00;                    // -> NaN/Inf
            else if( exp != 0 )                   // normalized value
            {
                exp += 0x1c000;                   // exp - 15 + 127
                if( mant == 0 && exp > 0x1c400 )  // smooth transition
                    return BitConverter.ToSingle(BitConverter.GetBytes( ( fullFloat & 0x8000 ) << 16
                                                    | exp << 13 | 0x3ff ), 0);
            }
            else if( mant != 0 )                  // && exp==0 -> subnormal
            {
                exp = 0x1c400;                    // make it normal
                do {
                    mant <<= 1;                   // mantissa * 2
                    exp -= 0x400;                 // decrease exp by 1
                } while( ( mant & 0x400 ) == 0 ); // while not normal
                mant &= 0x3ff;                    // discard subnormal bit
            }                                     // else +/-0 -> +/-0
            return BitConverter.ToSingle(BitConverter.GetBytes(          // combine all parts
                ( fullFloat & 0x8000 ) << 16          // sign  << ( 31 - 15 )
                | ( exp | mant ) << 13 ), 0);         // value << ( 23 - 10 )
        }
        /// <summary>
        /// Reads from the memory locations pointed to by the DataBox and saves it into a byte array
        /// ignoring the alpha component of each pixel.
        /// </summary>
        /// <param name="mapSource"></param>
        /// <returns></returns>
        private byte[] ToRGBArray(MappedSubresource mapSource, Format format)
        {
            byte[] bytes = new byte[CaptureWidth * 3 * CaptureHeight];
            int byteIndex = 0;

            if (format == Format.R16G16B16A16_Float)
            {
                unsafe
                {
                    byte* ptr = (byte*)mapSource.DataPointer;
                    byte* rowptr = ptr;
                    for (int y = 0; y < CaptureHeight; y++)
                    {
                        byte* pixelptr = rowptr;
                        for (int x = 0; x < CaptureWidth; x++)
                        {
                            for (int comp = 0; comp < 3; comp++)
                            {
                                byte lo = *pixelptr++;
                                byte hi = *pixelptr++;

                                // No idea why these values range from 4.6 to 0 instead of 1 to 0
                                // f(x) = sqrt(x/4.6) seems to approximate what the values should be.
                                bytes[byteIndex++] = (byte)(MathExt.Clamp(Math.Sqrt(Parse16BitFloat(hi, lo) / 4.6), 0, 1) * 255);
                            }
                            pixelptr += 2; //skip alpha
                        }
                        rowptr += mapSource.RowPitch;
                    }
                }
            }
            else if (format == Format.B8G8R8A8_UNorm)
            {
                IntPtr sourcePtr = mapSource.DataPointer;

                for ( int y = 0; y < CaptureHeight; y++ )
                {
                    Int32[] rowData = new Int32[CaptureWidth];
                    Marshal.Copy(sourcePtr, rowData, 0, CaptureWidth);

                    foreach ( Int32 pixelData in rowData )
                    {
                        byte[] values = BitConverter.GetBytes(pixelData);
                        if ( BitConverter.IsLittleEndian )
                        {
                            // Byte order : bgra
                            bytes[byteIndex++] = values[2];
                            bytes[byteIndex++] = values[1];
                            bytes[byteIndex++] = values[0];
                        }
                        else
                        {
                            // Byte order : argb
                            bytes[byteIndex++] = values[1];
                            bytes[byteIndex++] = values[2];
                            bytes[byteIndex++] = values[3];
                        }
                    }

                    sourcePtr = IntPtr.Add(sourcePtr, mapSource.RowPitch);
                }
            }
            else
            {
                throw new NotImplementedException($"Texture format {format.ToString()} is not supported.");
            }
            return bytes;
        }

        public void DelayNextCapture()
        {
            int remainingFrameTime = _minCaptureTime - (int)_captureTimer.ElapsedMilliseconds;
            if ( remainingFrameTime > 0 )
            {
                Thread.Sleep(remainingFrameTime);
            }
        }

        public void Dispose()
        {
            _duplicatedOutput?.Dispose();
            _output5?.Dispose();
            _output?.Dispose();
            _stagingTexture?.Dispose();
            _smallerTexture?.Dispose();
            _smallerTextureView?.Dispose();
            _device?.Dispose();
            _adapter?.Dispose();
            _factory?.Dispose();
            _lastCapturedFrame = null;
            _disposed = true;
        }

        public bool IsDisposed()
        {
            return _disposed;
        }
    }
}
