import {TestBed} from '@angular/core/testing';

import {ExcelReaderService} from './excel-reader.service';

describe('ExcelReaderService', () => {
  let service: ExcelReaderService;

  beforeEach(() => {
    TestBed.configureTestingModule({});
    service = TestBed.inject(ExcelReaderService);
  });

  it('should read items from the file', async () => {
    // Arrange
    const buffer = await fetch('/app/excel-reader/test-assets/test.spec.xlsx')
      .then((response: Response) => response.arrayBuffer());

    // Act
    await service.loadWorkbook(buffer);
    const items = service.readItems(1, 'A2:B11');

    // Assert
    expect(items).toEqual(['v21', 'v22', 'v31', 'v41', 'v42', 'v51', 'v61', 'v71', 'v91', 'v11_1']);
  });

  it('should auto determine range from the file', async () => {
    // Arrange
    const buffer = await fetch('/app/excel-reader/test-assets/test.spec.xlsx')
      .then((response: Response) => response.arrayBuffer());

    // Act
    await service.loadWorkbook(buffer);
    const range = service.autoDetermineRange();

    // Assert
    expect(range).toEqual('A1:A11');
  });
});
