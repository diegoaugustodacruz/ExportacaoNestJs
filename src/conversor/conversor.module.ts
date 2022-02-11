import { Module } from '@nestjs/common';
import { ConversorController } from './conversor.controller';
import { ConversorService } from './conversor.service';

@Module({
  controllers: [ConversorController],
  providers: [ConversorService],
})
export class ConversorModule {}
