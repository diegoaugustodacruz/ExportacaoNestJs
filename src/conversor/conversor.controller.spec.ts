import { Test, TestingModule } from '@nestjs/testing';
import { ConversorController } from './conversor.controller';

describe('ConversorController', () => {
  let controller: ConversorController;

  beforeEach(async () => {
    const module: TestingModule = await Test.createTestingModule({
      controllers: [ConversorController],
    }).compile();

    controller = module.get<ConversorController>(ConversorController);
  });

  it('should be defined', () => {
    expect(controller).toBeDefined();
  });
});
