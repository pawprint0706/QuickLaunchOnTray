using System;

namespace QuickLaunchOnTray.Models
{
    public class ProgramItem
    {
        public string Name { get; set; }  // ini 파일의 key 값 또는 단순 경로인 경우 파일명(확장자 제외)
        public string Path { get; set; }  // 프로그램 경로
    }
} 