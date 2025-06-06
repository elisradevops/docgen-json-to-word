﻿using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace JsonToWord.Services.Interfaces
{
    public interface IContentControlService
    {
        void ClearContentControl(WordprocessingDocument document, string contentControlTitle, bool force);
        SdtBlock FindContentControl(WordprocessingDocument preprocessingDocument, string contentControlTitle);
        void RemoveContentControl(WordprocessingDocument document, string contentControlTitle);
        void RemoveAllStdContentBlock(SdtBlock sdtBlock);
    }
}
