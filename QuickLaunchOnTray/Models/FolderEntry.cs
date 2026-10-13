using System.Collections.Generic;
using System.Drawing;

namespace QuickLaunchOnTray.Models
{
    /// <summary>폴더 메뉴를 구성하는 한 항목(파일 또는 하위 폴더).</summary>
    public class FolderEntry
    {
        public FolderEntry()
        {
            IsFolder = false;
            HasChildren = false;
        }

        /// <summary>표시 이름(파일명 또는 폴더명).</summary>
        public string Name { get; set; }

        /// <summary>전체 경로.</summary>
        public string Path { get; set; }

        /// <summary>하위 폴더 여부.</summary>
        public bool IsFolder { get; set; }

        /// <summary>하위 폴더일 때 자식 항목 존재 여부.</summary>
        public bool HasChildren { get; set; }

        /// <summary>메뉴 표시용 아이콘 이미지. 없으면 null(호출자가 기본 이미지를 사용).</summary>
        public Bitmap Image { get; set; }
    }
}
