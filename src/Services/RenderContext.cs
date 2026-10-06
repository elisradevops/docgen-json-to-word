using System;
using System.Threading;

namespace JsonToWord.Services
{
    /// <summary>
    /// What the render in progress on this request owns. FileService is a singleton, shared by every render,
    /// so anything specific to one request (here: where its attachments are staged) cannot live in a field of
    /// it; WordService.Create sets it for the duration of its (synchronous) render and FileService reads it.
    /// Outside a render (a unit test, a direct call) it is null and the default folder is used.
    /// </summary>
    public static class RenderContext
    {
        private static readonly AsyncLocal<string> CurrentAttachmentsFolder = new AsyncLocal<string>();

        public static string AttachmentsFolder => CurrentAttachmentsFolder.Value;

        public static IDisposable UseAttachmentsFolder(string folder)
        {
            var previous = CurrentAttachmentsFolder.Value;
            CurrentAttachmentsFolder.Value = folder;
            return new Restore(() => CurrentAttachmentsFolder.Value = previous);
        }

        private sealed class Restore : IDisposable
        {
            private readonly Action _restore;
            public Restore(Action restore) { _restore = restore; }
            public void Dispose() => _restore();
        }
    }
}
