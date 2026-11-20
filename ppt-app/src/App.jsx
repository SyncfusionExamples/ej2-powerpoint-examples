import {
  Presentation, SlideLayoutType, 
  HorizontalAlignmentType, TextSpacingType 
 } from '@syncfusion/ej2-pptx'

function App() {
  
  const IMAGE_URL1 = 'src/assets/cycleLogo.jpg';
  const IMAGE_URL2 = 'src/assets/adventure.png';

  const createAndDownloadPresentation = async (
    fileName = 'Output.pptx'
  ) => {
    const presentation = Presentation.create();

    const cycleLogo = await fetchBytes(IMAGE_URL1);
    const adventure = await fetchBytes(IMAGE_URL2);

    const firstSlide = presentation.slides.add(
      { layout: SlideLayoutType.Blank }
    );
    const titleBox1 = firstSlide.shapes.addTextBox(
      { bounds: 
        { x: 55, y: 23, width: 853, height: 72 },
      }
    );

    const titlePara1 = titleBox1.textBody.addParagraph();
    titlePara1.horizontalAlignment = HorizontalAlignmentType.Center;
    const titleText1 = titlePara1.addTextPart('Adventure Works Cycles');
    titleText1.font.fontSize = 36;

    const contentBox = firstSlide.shapes.addTextBox(
      { bounds:
        { x: 66, y: 132, width: 543, height: 350 },
      }
    );

    const para1 = contentBox.textBody.addParagraph();
    para1.firstLineIndent = -22;
    para1.lineSpacing = { type: TextSpacingType.Points, value: 38 };
    para1.spaceBefore = { type: TextSpacingType.Points, value: 10 };
    para1.addTextPart(
      'Adventure Works Cycles, the fictitious company on which the Adventure Works sample databases are based, is a large, multinational manufacturing company.'
    );

    const para2 = contentBox.textBody.addParagraph();
    para2.firstLineIndent = -22;
    para2.lineSpacing = { type: TextSpacingType.Points, value: 38 };
    para2.spaceBefore = { type: TextSpacingType.Points, value: 10 };
    para2.addTextPart(
      'The company manufactures and sells metal and composite bicycles to North American, European, and Asian commercial markets. While its base operation is located in Bothell, Washington, with 290 employees, several regional sales teams are located throughout their market base.'
    );
    firstSlide.shapes.addPicture(
      { data: cycleLogo, 
        bounds: 
        {
          x: 620, y: 246, width: 328, height: 123
        }
      }
    );

    const secondSlide = presentation.slides.add({
      layout: SlideLayoutType.Blank
    });
    const titleBox2 = secondSlide.shapes.addTextBox(
      { bounds: 
        { x: 62, y: 22, width: 830, height: 66 }
      }
    );

    const titlePara2 = titleBox2.textBody.addParagraph();
    titlePara2.horizontalAlignment = HorizontalAlignmentType.Center;
    const titleText2 = titlePara2.addTextPart('About Adventure Works Cycles');
    titleText2.font.fontName = 'Calibri';
    titleText2.font.fontSize = 40;

    secondSlide.shapes.addPicture(
      { data: adventure,
        bounds: { x: 172, y: 155, width: 643, height: 256}
      }
    );

    const bytes = await presentation.save();
    const blob = new Blob([bytes], {
      type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
    });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = fileName;
    anchor.click();
    URL.revokeObjectURL(url);

  }

  const fetchBytes = async (url) => {
    const response = await fetch(url, { cache: "no-cache" });
    console.log(response);
    if (!response.ok) {
      throw new Error(
        `Failed to load ${url}: ${response.status} ${response.statusText}`
      );
    }
    return new Uint8Array(await response.arrayBuffer());
  };

  return (
    <>
      <div
        style={{
          minHeight: '100svh',
          boxSizing: 'border-box',
          padding: 24,
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          backgroundColor: '#edf5ff',
        }}
      >
        <div
          style={{
            boxSizing: 'border-box',
            width: 'min(560px, 100%)',
            padding: '32px 30px',
            display: 'flex',
            flexDirection: 'column',
            gap: 14,
            backgroundColor: '#fff',
            borderRadius: 24,
            boxShadow: '0 18px 45px rgba(41, 75, 114, 0.12)',
            textAlign: 'center',
          }}
        >
          <p style={{ margin: 0, color: '#52677e' }}>
            Click the button to view a PowerPoint presentation generated. Please
            note that Microsoft PowerPoint viewer or its equivalent is required
            to view the resultant document.
          </p>
          <button
            style={{
              minHeight: 48,
              border: 0,
              borderRadius: 10,
              backgroundColor: '#126cbd',
              color: '#fff',
              font: 'inherit',
              fontWeight: 700,
              cursor: 'pointer',
            }}
            onClick={() =>
              createAndDownloadPresentation(
                "CreatePresentation.pptx"
              )
            }
          >
            Generate PPT
          </button>
        </div>
      </div>
    </>
  )
}

export default App
