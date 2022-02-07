import { Test, TestingModule } from '@nestjs/testing';
import { ConversorService } from './conversor.service';

describe('ConversorService', () => {
  let service: ConversorService;

  beforeEach(async () => {
    const module: TestingModule = await Test.createTestingModule({
      providers: [ConversorService],
    }).compile();

    service = module.get<ConversorService>(ConversorService);
  });

  it('should be defined', () => {
    expect(service).toBeDefined();
  });
});
